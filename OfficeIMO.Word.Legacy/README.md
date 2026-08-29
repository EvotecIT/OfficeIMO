# OfficeIMO.Word.Legacy

`OfficeIMO.Word.Legacy` safely reads selected WordPerfect, WordStar, Ami Pro, Lotus Word Pro, Microsoft Works/Write, and Word for DOS sources into the normal `OfficeIMO.Word.WordDocument` model.

```csharp
using OfficeIMO.Word.Legacy;

using LegacyWordImportResult imported = LegacyWordImporter.Import("archive.wpd");
Console.WriteLine(imported.Report.Quality);
foreach (OfficeCompatibilityFinding finding in imported.Report.Findings) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}

imported.Document.Save("archive.docx");
```

The importer is deliberately read-only. It never saves back to a legacy format, executes macros or embedded code, activates embedded objects, or resolves external links. Each result states whether recovery was structured or salvage quality and includes explicit feature-level loss diagnostics. The existing OfficeIMO Word converter packages can export the returned document to ODT, HTML, Markdown, or PDF; `PlainText` provides the bounded text view directly.

Detection combines stable file signatures with an optional source name. Use `FormatHint` only for a known damaged or weakly identified source. Resource limits and cancellation are enforced before and during parsing.
