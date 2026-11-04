# OfficeIMO.OneNote.Html

`OfficeIMO.OneNote.Html` converts the typed offline OneNote model to HTML without Microsoft Graph, a OneNote installation, or a commercial dependency. Choose semantic HTML when document structure and normal text flow matter, or visual HTML when the positioned OneNote page should remain intact.

## Semantic HTML

The semantic path emits ordinary reflowable HTML through the first-party `OfficeIMO.Markdown` document shell. OneNote runs retain their font family, point size, foreground and highlight colors, bold, italic, underline, strikethrough, subscript, and superscript styling. Structural and source diagnostics continue to use the shared OneNote projection contracts:

```csharp
using OfficeIMO.Html;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;

OneNoteSection section = OneNoteSectionReader.Read("Section.one");
string html = section.ToHtml(null, out OfficeConversionReport report);
section.SaveAsHtml("Section.html");
```

Use `OneNoteMarkdownExportOptions` to include conflict copies or version-history pages and to resolve extracted asset destinations. HTML rendering remains fully offline unless a caller explicitly configures external assets in `MarkdownHtmlExportOptions`. The separate Markdown conversion remains syntax-oriented and reports font, color, underline, and script details that portable Markdown cannot represent; semantic HTML no longer inherits that loss.

## Visual page HTML

The visual path maps each page to the same `OfficeDrawing` canvas used by OneNote image and visual PDF export, then embeds it as responsive SVG. It retains freeform placement, tables, images and printouts, ink, and structured math without maintaining a second HTML renderer.

```csharp
var options = new OneNoteVisualHtmlExportOptions {
    DocumentTitle = "Project notebook",
    Language = "en",
    IncludeAccessibleText = true
};

string html = section.ToVisualHtml(options);
section.SaveAsVisualHtml("Section-visual.html", options);
```

Set `Fragment = true` on `OneNoteVisualHtmlExportOptions` for embeddable page figures without a document shell; semantic export uses `OneNoteHtmlExportOptions.Fragment` the same way. Visual export reports page-mapping and image-fallback diagnostics in its `OfficeConversionReport`. With `IncludeAccessibleText` enabled, each page also contains an encoded assistive text projection for indexing and screen-reader context. The SVG remains the visual surface; use semantic HTML when ordinary DOM text selection and reflow are the primary requirement.

Set `OneNoteVisualHtmlExportOptions.DiagnosticSink` to a caller-owned collection when image-codec fallbacks or page-mapping warnings must be audited. Unsupported source images remain visible as placeholders in the SVG instead of disappearing silently.

## Import ordinary HTML

The reverse adapter uses the same generic section and block projector as Excel and PowerPoint. Headings form page boundaries; paragraphs, inline formatting, links, lists, tables, and bounded data URI images become typed offline OneNote content.

```csharp
using OfficeIMO.Html;
using OfficeIMO.OneNote.Html;

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
OneNoteSection section = source.ToOneNoteSection(
    new HtmlOneNoteExportOptions { SectionName = "Imported notes" }, out OfficeConversionReport report);
```

Use `ToOneNoteNotebook()` to create a one-section notebook. Both throw when the conversion reports an error diagnostic; set `FailOnLoss` to reject any reported loss. `HtmlOneNoteExportOptions.Limits` applies the shared native import budgets before pages, elements, tables, or image bytes are allocated.

Native HTML and ARIA tables remain editable. Uneven rows and spans become a rectangular grid with blank covered cells and an approximation diagnostic. Authored table and cell widths are bounded to the native page layout; otherwise columns share the available width according to their content. Captions are retained as preceding paragraphs, including bounded links. Image hyperlinks are preserved.

Applicable external stylesheet links and unsupported `colgroup`/`col` widths produce loss diagnostics. Inline styles remain available to the semantic projector.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OneNote.Html` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
