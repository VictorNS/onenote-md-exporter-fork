# Fork changes

This document lists the differences between this fork and the original
[alxnbl/onenote-md-exporter](https://github.com/alxnbl/onenote-md-exporter) repository.

## Skip empty pages

**File:** `src/OneNoteMdExporter/Services/Export/ExportServiceBase.cs`

If a page's Markdown is empty or contains only whitespace after conversion, the
exporter does not write a `.md` file for it. A debug message is logged instead.

**Why:** Empty OneNote pages used to produce empty Markdown files that cluttered
the output.

## Page title as Markdown heading

**File:** `src/OneNoteMdExporter/Services/Export/MdExportService.cs`

Each exported `.md` file now starts with the page title as a level-1 heading
(`# Title`). If a front matter header is enabled, the heading follows it.

**Why:** File names are truncated to `MdMaxFileLength`, so long page titles are
lost in the file name. The heading keeps the full original title in the
document.

## Monospace text exported as Markdown code

**Files:** `src/OneNoteMdExporter/Services/Export/ExportServiceBase.cs`,
`src/OneNoteMdExporter/Services/ConverterService.cs`

Text set in a monospace font (`Consolas` or `Courier New`) in OneNote is exported
as Markdown code instead of plain text or an HTML `<span>`.

**Why:** In OneNote, code snippets are usually marked only by a monospace font.
The original exporter lost this formatting, or kept it as
`<span style='font-family:...'>`, which is noisy and not rendered as code by
Markdown viewers. Pandoc also changed the content of code. It escaped characters
(for example, `a>1` became `a\>1`) and collapsed indentation.

This feature only runs when `UseHtmlStyling` is enabled, the same as the other
HTML styling features.

### Code blocks

If the monospace font is set on the whole paragraph (the `style` attribute of the
parent OE element), the line is treated as a code line. Consecutive code lines
are merged into one fenced code block:

````markdown
```
var a=12;
var b=a+5;
```
````

- The fence has no language identifier, because the language cannot be
  detected reliably.
- Empty lines between code lines do not split the block.
- A background colour alone does not mark text as code. Only the font does.

### Inline code

If the monospace font is set on a `<span>` inside a paragraph, the fragment
becomes inline code:

```markdown
Some text `const az = "bla";` more text.
```

- Adjacent monospace fragments are merged into a single inline code span.
- Leading and trailing spaces are moved outside the backticks.
- If the code itself contains backticks, a longer fence is used (for example,
  ``` ``a`b`` ```), following CommonMark rules.
- Other styles on the same fragment, such as background, are dropped.

### Implementation notes

- **Before Pandoc** (`ExportServiceBase.EscapeStylingSpan`): the OneNote XML is
  scanned for monospace fonts. The plain text of each code fragment (HTML tags
  removed, entities decoded) is Base64-encoded and wrapped in a marker:
  - `«code:<base64>»` for a whole code line;
  - `«icode:<base64>»` for an inline fragment.

  Base64 contains only `A-Z a-z 0-9 + / =`, so Pandoc passes the content
  through unchanged: no escaping and no whitespace collapsing.
- **After Pandoc** (`ConverterService.PageMdPostConversion`):
  - `ConvertCodeLines` turns runs of `«code:…»` lines into fenced code blocks;
  - `ConvertInlineCode` turns `«icode:…»` markers into inline code.
- Font detection handles OneNote's quirks: font names written without quotes,
  with single quotes, or with double quotes (`'Courier New'`, `"Courier New"`),
  and style attributes broken across lines.
