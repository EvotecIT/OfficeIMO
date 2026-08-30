# OfficeIMO.Word.Legacy

`OfficeIMO.Word.Legacy` safely reads selected WordPerfect, WordStar, Ami Pro, Lotus Word Pro, Microsoft Works/Write, and Word for DOS sources into the normal `OfficeIMO.Word.WordDocument` model.

```csharp
using OfficeIMO.Word.Legacy;

using LegacyWordImportResult imported = LegacyWordImporter.Import("archive.wpd");
Console.WriteLine(imported.Report.Quality);
foreach (LegacyWordParagraphContent paragraph in imported.Content.Paragraphs) {
    Console.WriteLine($"{paragraph.StyleName}: {paragraph.Text}");
}
foreach (OfficeCompatibilityFinding finding in imported.Report.Findings) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}

imported.Document.Save("archive.docx");
```

The importer is deliberately read-only. It never saves back to a legacy format, executes macros or embedded code, activates embedded objects, or resolves external links. Each result states whether recovery was structured or salvage quality and includes explicit feature-level loss diagnostics. `Content` retains source-oriented paragraphs, formatted runs, notes, and inert resource references beside the projected document. The existing OfficeIMO Word converter packages can export the returned document to ODT, HTML, Markdown, or PDF; `PlainText` provides the bounded text view directly.

## Profile coverage

| Family/profile | Quality | Recovered today | Explicit boundary |
| --- | --- | --- | --- |
| WordStar 3-7 character streams | Structured | hard and soft returns, paragraphs, bold, italic, underline, strike, superscript, subscript, page breaks, selected dot commands, bounded notes/comments, paragraph-style names, and inert graphics references | printer/font/color/style-library sequences and unrecognized dot commands are reported; text-marker lists are identified as inferred |
| Ami Pro SAM 4 | Structured | `[tag]` style definitions, `[edoc]` paragraphs, basic character styles, fonts, RGB color, alignment, spacing, page-break and keep properties, and applied source style names | frames, equations, images, tables, and unrecognized inline tags are inert or reported rather than guessed |
| Weak WordStar family or non-SAM4 Ami Pro input with an explicit hint | Salvage | bounded text and paragraphs | the structured profile is not claimed without coherent WordStar control/paragraph/EOF grammar or an exact Ami Pro SAM 4 header |
| WordPerfect 5/6 | Salvage | bounded document-area text, paragraphs, document-area offset, and active-content marker inventory | prefix packets, formatting codes, notes, tables, graphics, and layout are not yet semantically decoded |
| Lotus Word Pro LWP | Salvage | bounded text plus compound-content safety inventory | document zones, styles, notes, tables, graphics, and layout are not yet reconstructed |
| Microsoft Works word 2-8 | Salvage | bounded text and paragraphs plus compound-content safety inventory where applicable | formatting, fields, notes, tables, images, and layout are not yet reconstructed |
| Microsoft Write WRI | Salvage | bounded text and paragraph runs | formatting runs, objects, headers, footers, and layout are not yet reconstructed |
| Microsoft Word for DOS 4-6 | Salvage | bounded text and paragraphs | formatting, annotations, objects, and layout are not yet reconstructed |

`Structured` means the input matched and passed the documented profile grammar. It does not mean lossless conversion: inspect `Report.Findings`, or call `Report.RequireStructuredNoLoss()` when the workflow must reject every known approximation.

Detection combines stable file signatures and validated profile grammar with an optional source name. Use `FormatHint` only for a known damaged or weakly identified source; a hint selects the family but does not upgrade weak input to structured quality. Resource limits, including formatted-run and tag inventories, and cancellation are enforced before and during parsing.
