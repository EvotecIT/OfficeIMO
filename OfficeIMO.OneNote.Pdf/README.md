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
