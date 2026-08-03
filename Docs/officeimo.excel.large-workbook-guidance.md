# OfficeIMO.Excel Large Workbook Guidance

This guide describes the current safe path for large workbook generation, reading, and edit workflows. It is intentionally scoped to features with source support, tests, and benchmark artifacts in this repository.

## Recommended Generation Paths

| Workload | Preferred API | Notes |
| --- | --- | --- |
| DataSet or DataTable export | `InsertDataSet(...)`, `InsertDataTable(...)`, then `Save(...)` | Fast package writers are selected automatically when the workbook shape is eligible. Use `ExcelSaveOptions.DisableFastPackageWriter` only for troubleshooting or comparative validation. |
| Object reports | `InsertObjects(...)`, table helpers, and one document-boundary save | Keep layout, AutoFit, tables, and formulas in one document session so shared strings, styles, and package finalization are batched. |
| Wide reports with AutoFit | `Execution.SaveWorksheetAfterAutoFit = false`, then `Save(...)` | Defers worksheet-part writes until the document boundary. This is the recommended report-export mode for large generated sheets. |
| Formula-backed reports | `doc.Calculate()` or `ExcelSaveOptions.EvaluateFormulasBeforeSave` | Only supported formula shapes are evaluated. Pair unsupported formulas with `ForceFullCalculationOnOpen` when the spreadsheet app should finish calculation. |

## Recommended Read Paths

| Workload | Preferred API | Notes |
| --- | --- | --- |
| Forward-only Excel reads | `ExcelDocument.OpenDataReader(...)` | The package-owned `DbDataReader` contract covers XLSX, XLSM, XLSB, and BIFF8 XLS and discovers used ranges automatically. |
| Forward-only CSV reads | `CsvDocument.OpenDataReader(...)` | The parallel API remains in `OfficeIMO.CSV`, with CSV-specific delimiter, encoding, compression, and schema options. |
| Multiple worksheets | `DbDataReader.NextResult()` | Results stay in workbook order. Set `ExcelReadOptions.SheetName` when only one worksheet should be opened. |
| Unknown workbook edit intake | `ExcelDocument.Load(...)`, `InspectFeatures()`, and `InspectFormulas()` | Use the editable document model only when the workbook will be inspected, mutated, converted, or saved again. Treat preserve-only and unsupported findings as a preflight signal. |
| Large Open XML edit intake | `ExcelDocument.OpenFileBacked(...)` | Stages the editable package in an owner-only temporary file, enforces `MaxInputBytes`, honors Open XML part limits, and observes cancellation while copying. XLS/XLSB projection continues through `Load`. |

Example:

```csharp
using OfficeIMO.Excel;

using var reader = ExcelDocument.OpenDataReader("sales.xlsx");
int revenue = reader.GetOrdinal("Revenue");
while (reader.Read()) {
    Console.WriteLine(reader.GetDecimal(revenue));
}
```

## Preflight Before Editing Existing Workbooks

Run feature inspection before mutating workbooks that were not created by the current workflow:

```csharp
using var document = ExcelDocument.Load(path);
ExcelFeatureReport features = document.InspectFeatures();
features.EnsureNoUnsupportedFeatures();

foreach (ExcelFeatureFinding feature in features.PreservedFeatures) {
    Console.WriteLine($"{feature.Name}: {feature.Count}");
}
```

For formula-heavy files, inspect formula support separately:

```csharp
ExcelFormulaInspection formulas = document.InspectFormulas();
Console.WriteLine(formulas.Capabilities.Summary);
formulas.EnsureAllHaveCachedResults();
```

Use `EnsureNoAdvancedFeatures()` only for workflows that must avoid preserve-only package content such as custom XML, macros, slicers, timelines, embedded packages, or external workbook relationships.

For workflow routing, prefer the capability preflight API over ad hoc feature-name checks:

```csharp
ExcelFeatureReport features = document.InspectFeatures();

features.EnsureCan(ExcelPreflightCapability.EditCellValues);

if (!features.Can(ExcelPreflightCapability.ExportPdfReport)) {
    File.WriteAllText("excel-preflight.md", features.ToMarkdown());
}
```

`ExcelPreflightCapability` covers readback, cell-value edits, structure-changing edits, cached formula reads, OfficeIMO formula calculation, template binding, and first-party PDF report export. This is intentionally separate from benchmark guidance and does not require benchmark runs in CI.

For structure-changing work, produce a bounded plan before committing:

```csharp
var limits = new ExcelMutationPlanOptions {
    MaximumAffectedCells = 500_000,
    MaximumSnapshotCharacters = 64_000_000,
    MaximumDiagnostics = 50
};

ExcelStructuralMutationPlan plan = document["Data"].PlanMoveRange("A2:F50000", "H2", limits);
ExcelMutationResult result = plan.Apply(cancellationToken);
```

The transaction revalidates ownership conflicts, retains a bounded rollback snapshot, and returns bounded post-edit Open XML diagnostics. Established direct writers, streaming readers, ordinary `Load`, and unchanged-package save paths do not route through this transaction unless the caller selects it.

## Measuring A Change

Use the benchmark harness for repeatable local evidence:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --skip-legacy-epplus --warmup 20 --iterations 9
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- write-profile --rows 25000
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- read-profile --rows 25000
```

Use the comparison summary for public-facing numbers only when the run records:

- row counts and scenario names
- Release configuration and target framework
- raw samples, mean, median, and allocation data
- package-size and package-part metrics when save behavior matters
- machine and runtime information from the artifact manifest

The CSV and Excel benchmark projects own their respective library comparisons:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.CSV.Benchmarks -- --filter "*CsvBenchmarks*"
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks -- comparison-suite --out-dir .\artifacts\excel --row-set 2500,25000
```

The suites compare OfficeIMO with the configured libraries that support each equivalent workload. Product guidance uses the validated OfficeIMO contract; competitor-specific results stay in benchmark artifacts and the benchmark website.
Windows, Linux, and macOS remain separate evidence lanes; never average them or
substitute one platform when another platform is missing.

## Current Boundaries

- Large workbook guidance is strongest for generated report-style workbooks, bounded reads, and explicitly file-backed Open XML edits.
- Feature-rich externally authored workbooks should be inspected before mutation because preserve-only package parts may need round-trip care.
- Fast package writers are automatic optimizations, not a compatibility promise for every workbook shape.
- Rendering/export is not part of the current large-workbook promise.

## Related Evidence

- Benchmark artifact guide: `Docs/benchmarks/README.md`
- Benchmark notes: `Docs/officeimo.excel.benchmark-notes.md`
- Current capability matrix: `OfficeIMO.Excel/COMPATIBILITY.md`
