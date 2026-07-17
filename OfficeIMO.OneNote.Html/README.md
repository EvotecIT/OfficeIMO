# OfficeIMO.OneNote.Html

`OfficeIMO.OneNote.Html` converts the typed offline OneNote model to HTML without Microsoft Graph, a OneNote installation, or a commercial dependency. Choose semantic HTML when document structure and normal text flow matter, or visual HTML when the positioned OneNote page should remain intact.

## Semantic HTML

The semantic path uses `OfficeIMO.OneNote.Markdown` and the first-party `OfficeIMO.Markdown` HTML renderer:

```csharp
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;

OneNoteSection section = OneNoteSectionReader.Read("Section.one");
string html = section.ToHtmlDocument();
section.SaveAsHtml("Section.html");
```

Use `OneNoteMarkdownOptions` to include conflict copies or version-history pages and to resolve extracted asset destinations. HTML rendering remains fully offline unless a caller explicitly configures external assets in `HtmlOptions`.

## Visual page HTML

The visual path maps each page to the same `OfficeDrawing` canvas used by OneNote image and visual PDF export, then embeds it as responsive SVG. It retains freeform placement, tables, images and printouts, ink, and structured math without maintaining a second HTML renderer.

```csharp
var options = new OneNoteVisualHtmlOptions {
    DocumentTitle = "Project notebook",
    Language = "en",
    IncludeAccessibleText = true
};

string html = section.ToVisualHtmlDocument(options);
section.SaveAsVisualHtml("Section-visual.html", options);
```

`ToVisualHtmlFragment(...)` returns embeddable page figures without a document shell. With `IncludeAccessibleText` enabled, each page also contains an encoded assistive text projection for indexing and screen-reader context. The SVG remains the visual surface; use semantic HTML when ordinary DOM text selection and reflow are the primary requirement.

Set `OneNoteVisualHtmlOptions.DiagnosticSink` to a caller-owned collection when image-codec fallbacks or page-mapping warnings must be audited. Unsupported source images remain visible as placeholders in the SVG instead of disappearing silently.
