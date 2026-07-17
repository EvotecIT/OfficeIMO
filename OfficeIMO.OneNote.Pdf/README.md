# OfficeIMO.OneNote.Pdf

`OfficeIMO.OneNote.Pdf` converts the typed offline OneNote model to PDF without Microsoft Graph, a OneNote installation, or a commercial dependency. It projects OneNote once through `OfficeIMO.OneNote.Markdown`, then uses the first-party `OfficeIMO.Markdown.Pdf` and `OfficeIMO.Pdf` engines.

```csharp
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Pdf;

OneNoteSection section = OneNoteSectionReader.Read("Section.one");
byte[] pdf = section.ToPdf();
section.SaveAsPdf("Section.pdf");
```

Use `OneNoteMarkdownOptions` for conflict/version inclusion and asset destinations, and `MarkdownPdfSaveOptions` for PDF layout, fonts, image policy, and diagnostics.

OneNote PDF export enables the shared multilingual system-font fallback in addition to the normal document, monospace, and symbol fallbacks. The converter clones caller options, so this compatibility default does not mutate a reusable `MarkdownPdfSaveOptions` instance. Actual glyph coverage depends on fonts installed on the host. Applications that require strict standard-font output can opt out explicitly:

```csharp
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;

var pdfOptions = new MarkdownPdfSaveOptions {
    TextFallbacks = PdfTextFallbackFeatures.None
};

section.SaveAsPdf("Section-standard-fonts.pdf", pdfOptions: pdfOptions);
```

Invalid Unicode noncharacters and native control codes are normalized by `OfficeIMO.OneNote.Markdown` before layout. Valid characters that remain uncovered by the configured fonts still produce the normal PDF conversion diagnostic rather than being silently discarded.
