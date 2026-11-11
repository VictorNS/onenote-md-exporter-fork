using alxnbl.OneNoteMdExporter.Infrastructure;
using alxnbl.OneNoteMdExporter.Models;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using System.Linq;

namespace alxnbl.OneNoteMdExporter.Services
{
    public static class ConverterService
    {

        /// <summary>
        /// Convert DocX file into MD using PanDoc
        /// </summary>
        /// <param name="page">OneNote  page to convert</param>
        /// <param name="inputFilePath">docx file exported from OneNote page</param>
        /// <param name="tmpFolderPath">path to the folder container attachments</param>
        public static string ConvertDocxToMd(Page page, string inputFilePath, string tmpFolderPath)
        {
            var tmpDir = Path.Combine(tmpFolderPath, "pandoc");
            if (Directory.Exists(tmpDir))
                Directory.Delete(tmpDir, true);
            Directory.CreateDirectory(tmpDir);

            var pandocPath = @"pandoc\pandoc.exe";
            if (!File.Exists(pandocPath))
                pandocPath = "pandoc.exe";

            var mdFilePath = Path.Combine(tmpDir, page.TitleWithNoInvalidChars(AppSettings.MdMaxFileLength) + ".md");
 
            var arguments = $"\"{Path.GetFullPath(inputFilePath)}\" " +
                            $"--to={AppSettings.PanDocMarkdownFormat} " +
                            $"-o \"{Path.GetFullPath(mdFilePath)}\" " +
                            $"--wrap=none " + // Mandatory to avoid random quote block to be added to markdown
                            $"--extract-media=\"{tmpDir}\"";

            var startInfo = new ProcessStartInfo
            {
                FileName = pandocPath,
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            Log.Debug($"{page.Id} : Start Pandoc");

            try
            {
                using Process exeProcess = Process.Start(startInfo);
                exeProcess.WaitForExit();

                if (exeProcess.ExitCode == 0)
                {
                    Log.Debug($"{page.Id} : Pandoc success");


                    if (AppSettings.Debug)
                        Log.Debug($"Pandoc output: {exeProcess.StandardOutput.ReadToEnd()}");

                    File.Delete(inputFilePath);

                    var mdFileContent = File.ReadAllText(mdFilePath);

                    return mdFileContent;
                }
                else
                {
                    Log.Error($"{page.Id} : Pandoc error");
                    var pandocError = exeProcess.StandardError.ReadToEnd();
                    Log.Error("Pandoc error output: {0}", pandocError);

                    throw new Exception("Error during PanDoc execution");
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
                throw new Exception(@"pandoc.exe not found in \pandoc\ subfolder. Have you unzip the pandoc archive ?");
            }
        }

        /// <summary>
        /// Apply post-conversion to md file
        /// </summary>
        /// <param name="page">Section page</param>
        /// <param name="mdFileContent"></param>
        public static void PageMdPostConversion(ref string mdFileContent)
        {
            if (AppSettings.DeduplicateLinebreaks)
            {
                mdFileContent = DeduplicateLinebreaks(mdFileContent);
            }

            mdFileContent = RemoveHtmlCommentBlocks(mdFileContent);

            mdFileContent = RemoveUTF8NonBreakingSpace(mdFileContent);

            if (AppSettings.PostProcessingRemoveQuotationBlocks)
            {
                mdFileContent = RemoveQuotationBlocks(mdFileContent);
            }

            if (AppSettings.MaxTwoLineBreaksInARow)
            {
                mdFileContent = MaxTwoLineBreaksInARow(mdFileContent);
            }

            if (AppSettings.PostProcessingRemoveOneNoteHeader)
            {
                mdFileContent = RemoveOneNoteHeader(mdFileContent);
            }

            if (AppSettings.UseHtmlStyling)
            {
                mdFileContent = ConvertCodeLines(mdFileContent);
                mdFileContent = ConvertInlineCode(mdFileContent);
                mdFileContent = UnEscapeStylingSpan(mdFileContent);
            }

            mdFileContent = InsertMdHighlight(mdFileContent);

            mdFileContent = InsertMdTodoMarks(mdFileContent);
        }

        private static string InsertMdTodoMarks(string mdFileContent)
        {
            return mdFileContent.Replace("«☐»", "- [ ] ").Replace("«☑»", "- [x] ");
           
        }

        private static string RemoveOneNoteHeader(string pageTxt)
        {
            var pageTxtModified = Regex.Replace(pageTxt, @"^.+(\n|\r|\r\n){1,2}.+(\n|\r|\r\n){1,2}\d{2}:\d{2}\s+", "");

            return pageTxtModified;
        }

        private static string RemoveUTF8NonBreakingSpace(string pageTxt)
        {
            // Max 2 consecutive linebreaks
            var pageTxtModified = Regex.Replace(pageTxt, @"(\xa0|\xc2|\xc2\xa0)", string.Empty);

            return pageTxtModified;
        }

        private static string RemoveHtmlCommentBlocks(string pageTxt)
        {
            // Pandoc produce <!-- --> tags
            var pageTxtModified = Regex.Replace(pageTxt, @"(\n|\r|\r\n)( )*\<!--( )*--\>( )*", "");

            return pageTxtModified;
        }

        

        private static string DeduplicateLinebreaks(string pageTxt)
        {
            // PanDoc seems to produce 2 linebreaks characters for each linebreak in original DocX file
            // Replace all pair of linebreak by a single linebreak
            var pageTxtModified = Regex.Replace(pageTxt, @"(\n>?\n|\r>?\r|\r\n>?\r\n)", Environment.NewLine);

            return pageTxtModified;
        }

        private static string MaxTwoLineBreaksInARow(string pageTxt)
        {
            // Max 2 consecutive linebreaks
            var pageTxtModified = Regex.Replace(pageTxt, @"((\n[ \t]*\n+)|(\r[ \t]*\r+)|(\r\n[ \t]*(\r\n)+))",
                Environment.NewLine + Environment.NewLine, RegexOptions.Multiline);

            return pageTxtModified;
        }

        private static string RemoveQuotationBlocks(string pageTxt)
        {
            string regex = @"(\n|\r|\r\n)>(\n|\r|\r\n)";
            var pageTxtModified = Regex.Replace(pageTxt, regex, delegate (Match match)
            {
                return Environment.NewLine;
            });

            string regex2 = @"(\n|\r|\r\n)>[ ]?";
            pageTxtModified = Regex.Replace(pageTxtModified, regex2, delegate (Match match)
            {
                return Environment.NewLine;
            });


            return pageTxtModified;
        }

        /// <summary>
        /// Replace PanDoc html tags <span class="mark">text</span> by ==text== 
        /// </summary>
        /// <param name="pageTxt"></param>
        /// <returns></returns>
        private static string InsertMdHighlight(string pageTxt)
        {
            // match and replace each span block of a row
            string regex = @"\<span class=\""mark\""\>(?<text>((?!\</span\>).)*)\</span\>"; // https://stackoverflow.com/questions/406230/regular-expression-to-match-a-line-that-doesnt-contain-a-word
            var pageTxtModified = Regex.Replace(pageTxt, regex, delegate (Match match)
            {
                return "==" + (match.Groups["text"]?.Value ?? "") + "==";
            });

            return pageTxtModified;
        }

        public const string CodeLineStartMarker = "«code:";
        public const string CodeLineEndMarker = "»";

        public const string InlineCodeStartMarker = "«icode:";

        /// <summary>
        /// Decode a full code marker («code:...» or «icode:...»)
        /// </summary>
        public static string DecodeCodeMarker(string marker)
        {
            var start = marker.IndexOf(':') + 1;
            return DecodeCodeLine(marker[start..^CodeLineEndMarker.Length]);
        }

        /// <summary>
        /// Replace inline code markers (monospace font fragments) by markdown inline code
        /// </summary>
        private static string ConvertInlineCode(string pageTxt)
        {
            return Regex.Replace(pageTxt, @"«icode:([A-Za-z0-9+/=]*)»", match =>
            {
                var text = DecodeCodeLine(match.Groups[1].Value);
                var code = text.Trim();
                if (code.Length == 0)
                    return text;

                var leading = text[..text.IndexOf(code[0])];
                var trailing = text[(leading.Length + code.Length)..];

                // Fence must be longer than the longest backtick run inside the code
                var maxBackticks = Regex.Matches(code, "`+").Select(m => m.Length).DefaultIfEmpty(0).Max();
                var fence = new string('`', maxBackticks + 1);
                var padding = code.StartsWith('`') || code.EndsWith('`') ? " " : "";

                return $"{leading}{fence}{padding}{code}{padding}{fence}{trailing}";
            });
        }

        /// <summary>
        /// Replace consecutive code line markers
        /// </summary>
        private static string ConvertCodeLines(string pageTxt)
        {
            var codeLineRegex = new Regex(@"^[ \t>]*«code:([A-Za-z0-9+/=]*)»\\?[ \t]*$");
            var inlineCodeRegex = new Regex(@"«code:([A-Za-z0-9+/=]*)»");

            var lines = Regex.Split(pageTxt, @"\r\n|\n|\r");
            var result = new List<string>();
            var codeBlock = new List<string>();
            var pendingBlankLines = 0;

            void FlushCodeBlock()
            {
                if (codeBlock.Count == 0)
                    return;
                result.Add("```");
                result.AddRange(codeBlock);
                result.Add("```");
                codeBlock.Clear();
            }

            foreach (var line in lines)
            {
                var match = codeLineRegex.Match(line);
                if (match.Success)
                {
                    // Blank lines between code lines are merged into the same block
                    pendingBlankLines = 0;
                    codeBlock.Add(DecodeCodeLine(match.Groups[1].Value));
                }
                else if (codeBlock.Count > 0 && string.IsNullOrWhiteSpace(line))
                {
                    pendingBlankLines++;
                }
                else
                {
                    FlushCodeBlock();
                    for (; pendingBlankLines > 0; pendingBlankLines--)
                        result.Add(string.Empty);
                    // Code marker in the middle of a text line: keep raw text
                    result.Add(inlineCodeRegex.Replace(line, m => DecodeCodeLine(m.Groups[1].Value)));
                }
            }

            FlushCodeBlock();
            for (; pendingBlankLines > 0; pendingBlankLines--)
                result.Add(string.Empty);

            return string.Join(Environment.NewLine, result);
        }

        private static string DecodeCodeLine(string base64)
        {
            return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(base64));
        }

        private static string UnEscapeStylingSpan(string pageTxt)
        {
            // Unescape the highlighting span tags (after Pandoc conversion)
            string highlightRegex = @"«span\s+style='(\s*[a-zA-Z0-9\s\.\#;:-]*)'»(.*?)«\/span»";
            var pageTxtModified = Regex.Replace(pageTxt, highlightRegex, delegate (Match match)
            {
                return $"<span style='{match.Groups[1]}'>{match.Groups[2]}</span>";
            });

            return pageTxtModified;
        }
    }
}
