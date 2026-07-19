# Migrating from OfficeIMO 2.x to 3.0

OfficeIMO 3.0 aligns the supported package set on one release line, makes table-only PDF recovery explicit, and removes public access to implementation details that applications should not have needed. Upgrade all OfficeIMO packages in an application together.

```xml
<PackageReference Include="OfficeIMO.Word" Version="3.0.0" />
<PackageReference Include="OfficeIMO.Excel" Version="3.0.0" />
<PackageReference Include="OfficeIMO.Pdf" Version="3.0.0" />
```

Do not mix 2.x and 3.0 OfficeIMO packages in the same dependency graph. Adapter packages reference their owning document and renderer packages from the same coordinated line.

## PDF table imports

The Excel and PowerPoint PDF adapters recover detected logical tables. They do not reproduce every PDF page element. Their 3.0 names now state that contract.

| OfficeIMO 2.x | OfficeIMO 3.0 |
|---|---|
| `SaveAsExcel` / `SaveAsExcelAsync` | `SaveTablesAsExcel` / `SaveTablesAsExcelAsync` |
| `ToExcelDocument` | `ImportTablesToExcelDocument` |
| `ToExcelDocumentResult` | `ImportTablesToExcelDocumentResult` |
| `PdfExcelConversionReport` | `PdfExcelTableImportReport` |
| `PdfExcelConversionResult` | `PdfExcelTableImportResult` |
| `SaveAsPowerPoint` / `SaveAsPowerPointAsync` | `SaveTablesAsPowerPoint` / `SaveTablesAsPowerPointAsync` |
| `ToPowerPointPresentation` | `ImportTablesToPowerPointPresentation` |
| `ToPowerPointPresentationResult` | `ImportTablesToPowerPointPresentationResult` |
| `PdfPowerPointConversionReport` | `PdfPowerPointTableImportReport` |
| `PdfPowerPointConversionResult` | `PdfPowerPointTableImportResult` |

Load a logical PDF, then call the table-specific adapter:

```csharp
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;

PdfLogicalDocument source = PdfLogicalDocument.Load("report.pdf");
PdfExcelTableImportReport report = source.SaveTablesAsExcel("tables.xlsx");

if (report.HasOmittedPageContent) {
    Console.WriteLine("The source also contains content outside detected tables.");
}
```

`HasLoss` means a detected table was truncated by an import limit. `HasOmittedPageContent` means the source also contains non-table text, vector graphics, images, links, forms, annotations, or page actions that the table-only adapter does not import. Use `SourceScope` for the counts behind that decision. Use Word/RTF semantic conversion or image rendering when a full-page representation is the goal.

## Word public surface

Several helper types were implementation details rather than stable application APIs:

| OfficeIMO 2.x | OfficeIMO 3.0 |
|---|---|
| `FormattingHelper.GetFormattedRuns(paragraph)` | `paragraph.GetFormattedRuns()` returning `WordFormattedRun` values |
| `WordListLevel._level` | `WordListLevel.OpenXmlElement` |
| `new WordHelpers()` | Remove the instance; `WordHelpers` is static and its supported methods are called directly |
| `WordHelpers.GetNextSdtId(...)` | Removed; content-control APIs allocate valid IDs internally |
| `InlineRunHelper.AddInlineRuns(...)` | Use the owning converter or explicit paragraph APIs |
| `ImageShapeStyleHelper` | Use the owning image shape APIs |
| `HorizontalAlignmentHelper` | Use the public alignment properties on the owning paragraph, table, cell, or image API |

For Markdown, parse the document through `OfficeIMO.Word.Markdown` instead of using the old inline-run helper:

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using WordDocument document = MarkdownReader.Parse(markdown).ToWordDocument();
```

`ConvertDotxToDocx` now resolves relative template paths before constructing the package URI, so relative and absolute template paths follow the same behavior.

## Legacy XLS import reports

The supported replacements are:

| OfficeIMO 2.x | OfficeIMO 3.0 |
|---|---|
| `LegacyXlsLoadResult.Workbook` | `LegacyXlsLoadResult.AdvancedWorkbook` |
| `LegacyXlsLoadResult.ImportReport` | `LegacyXlsLoadResult.CreateImportReport()` |
| `LegacyXlsLoadResult.CreateAdvancedImportReport()` | `LegacyXlsLoadResult.CreateImportReport()` |
| Detailed `LegacyXlsImportReport` record-family counters | Stable summary counts and issue collections |

`AdvancedWorkbook` is the public imported workbook. The low-level `Workbook` projection and exhaustive parser telemetry are internal in 3.0. `CreateImportReport()` returns the cached public report with the stable summary counts and the derived `HasImportErrors` and `HasUnsupportedFeatures` indicators. Detailed record-family counters remain available to OfficeIMO's import implementation and tests without becoming permanent public API.

## EPUB image export package

The EPUB-to-image adapter is now named for its result:

```text
OfficeIMO.Epub.Html  ->  OfficeIMO.Epub.Image
```

Update both the package reference and namespace imports. The adapter still retains EPUB chapter HTML and package resources internally and renders through the shared HTML image pipeline; the rename does not introduce another HTML renderer.

## Compatibility shim visibility

`OfficeIMO.Drawing` no longer exports `System.Runtime.CompilerServices.IsExternalInit` from its `netstandard2.0` and `net472` assets. That type was a compiler compatibility shim, not an OfficeIMO API. OfficeIMO still supplies an internal shim where the target framework needs one, so record and `init` usage in applications is unaffected. Remove any direct reference to the OfficeIMO-provided shim.

## Package and dependency ownership

OfficeIMO 3.0 keeps format ownership in the existing document, renderer, and adapter projects. There is no new catch-all core package. Small adapter packages such as PDF or image exporters remain thin surfaces over the owning parser and renderer, which avoids duplicating conversion logic or forcing unrelated dependencies into document packages.

After upgrading, perform a clean restore so lock files and cached transitive packages no longer retain 2.x OfficeIMO versions.

The [3.0 public API review](officeimo-3.0-public-api-review.md) records the complete assembly-level comparison used to confirm that these are the only changed coordinated public surfaces.

For the older 1.x to 2.0 API removals, see the [2.0 breaking API migration guide](officeimo.breaking-api-migration.md).
