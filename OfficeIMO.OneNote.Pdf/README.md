# OfficeIMO.OneNote.Pdf

`OfficeIMO.OneNote.Pdf` converts the typed offline OneNote model to a semantic PDF document without Microsoft Graph, a OneNote installation, or a commercial dependency. It projects OneNote once through `OfficeIMO.OneNote.Markdown`, then uses the first-party `OfficeIMO.Markdown.Pdf` and `OfficeIMO.Pdf` engines.

```csharp
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Pdf;

OneNoteSection section = OneNoteSectionReader.Read("Section.one");
byte[] pdf = section.ToPdf();
section.SaveAsPdf("Section.pdf");
```

OneNote pages are free-form canvases. The current `SemanticDocument` mode intentionally flattens them into reading order. `ToPdfDocumentResult()` reports canvas flattening, formatting simplification, unresolved asset placeholders, link-only binary assets, opaque omissions, and source diagnostics. It does not claim pixel parity with the OneNote desktop canvas.

Use `OneNotePdfSaveOptions.ProjectionOptions` for conflict/version inclusion and asset destinations, and `OneNotePdfSaveOptions.PdfOptions` for PDF layout, fonts, image policy, and diagnostics.

OneNote PDF export adds multilingual fallback candidates in addition to the normal document, monospace, and symbol candidates. The balanced default uses installed fonts while denying arbitrary local and remote reads; portable deterministic mode is explicit. Conversion clones both projection and PDF options so reusable caller configuration is not mutated:

```csharp
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;

var options = new OneNotePdfSaveOptions {
    PdfOptions = new MarkdownPdfSaveOptions {
        ResourcePolicy = PdfResourcePolicy.CreateTrustedHost()
    }
};

PdfDocumentConversionResult result = section.ToPdfDocumentResult(options);
foreach (PdfConversionWarning warning in result.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
result.Save("Section.pdf");
```

Invalid Unicode noncharacters and native control codes are normalized by `OfficeIMO.OneNote.Markdown` before layout. Valid characters that remain uncovered by the configured fonts still produce the normal PDF conversion diagnostic rather than being silently discarded.
