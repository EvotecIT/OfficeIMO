# OfficeIMO.Excel.Legacy

`OfficeIMO.Excel.Legacy` safely reads selected Lotus 1-2-3, Quattro Pro, Multiplan, and Microsoft Works spreadsheet sources into the normal `OfficeIMO.Excel.ExcelDocument` model.

## Install from an OfficeIMO source checkout

```powershell
dotnet add .\YourApp.csproj reference .\OfficeIMO.Excel.Legacy\OfficeIMO.Excel.Legacy.csproj
```

Keep the legacy importer and the other OfficeIMO projects in the same coordinated source revision.

```csharp
using OfficeIMO;
using OfficeIMO.Excel.Legacy;

using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import("archive.wk1");
Console.WriteLine(imported.Report.Quality);
foreach (LegacySpreadsheetCellContent cell in imported.Cells) {
    Console.WriteLine($"{cell.SheetName}!R{cell.Row}C{cell.Column}: {cell.Formula ?? cell.CachedValue}");
}
foreach (OfficeCompatibilityFinding finding in imported.Report.Findings) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}

imported.Document.Save("archive.xlsx");
```

The importer is deliberately read-only. It never saves back to a legacy format, executes macros, activates embedded objects, or resolves and refreshes external links. Each result states whether recovery was structured or salvage quality and includes explicit feature-level loss diagnostics. The existing OfficeIMO Excel converter packages can export the returned workbook to ODS, CSV, HTML, or PDF.

## Profile coverage

| Family/profile | Quality | Recovered today | Explicit boundary |
| --- | --- | --- | --- |
| Lotus 1-2-3 WK1 `0x0406` record streams | Structured | empty workbooks, blank and populated cells, labels, integers, doubles, finite formula caches, safely translated bounded RPN formulas, range names with 16-bit columns, selected number formats, label alignment, and chart-record metadata | WK1 is projected as one source sheet; other Lotus WK/123 BOF profiles remain salvage; unsupported formula tokens retain the cache with a diagnostic; every unprojected record kind is inventoried as loss; advanced styles, comments, and live chart reconstruction remain open |
| Quattro Pro DOS WQ1/WQ2 record streams | Structured | sheet identifiers, cells, finite cached values, range names, label alignment, and chart metadata | the bounded WQ2 profile accepts only zero values in its two unmodeled cell-header attribute bytes; the Quattro formula/reference dialect retains cached values with a diagnostic; WB/QPW structures, comments, advanced formatting, and live charts are not claimed |
| Microsoft Works DOS WKS `0x0404` record streams | Structured | the shared early record contract: cells, cached values, safe formulas, range names, selected number formats, alignment, and chart metadata | later Works binary/compound structures and comments are not claimed |
| Later Lotus 123, Quattro QPW, and Works XLR/binary profiles | Salvage | bounded text/tabular runs and compound-content safety inventory where applicable | workbook structure, formulas, names, comments, advanced formatting, and charts are reported as unavailable |
| Microsoft Multiplan DOS 1-3 | Salvage | bounded text and tabular runs | cell zones, formulas, names, formats, comments, and charts are not yet semantically decoded |

The three structured WK-derived profiles accept a structurally valid BOF/EOF workbook with no cells. Their current text contract is ASCII-only: an undeclared extended byte is rejected instead of being silently replaced, while an affected formula retains only its finite cached value with a loss diagnostic. Producer-specific code-page support remains roadmap work.

Formula translation is allow-listed, expression-depth/node/character bounded, charged against the import-wide recovered-text budget, and never evaluates the source expression. The source-oriented `Cells` collection retains each cached value, alignment, and translated formula. `Names` retains validated fixed-layout source records and exposes `ProjectedName`; it is null when strict Excel-name validation or collision handling kept the source name as metadata instead of silently rewriting or overwriting it. Unsupported formula tokens are never guessed: the finite cached value is projected and the loss report records that decision.

`Structured` means the record stream passed the documented profile grammar. It does not mean lossless conversion: inspect `Report.Findings`, or call `Report.RequireStructuredNoLoss()` when every known approximation must fail the workflow.

Path, stream, and byte-array inputs share the same limits, cancellation, detection, and loss-report contracts.
