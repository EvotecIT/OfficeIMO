# OfficeIMO.Excel Large Workbook Guidance

This guide describes the current safe path for large workbook generation, reading, and edit workflows. It is intentionally scoped to features with source support, tests, and benchmark artifacts in this repository.

## Recommended Generation Paths

| Workload | Preferred API | Notes |
| --- | --- | --- |
| DataSet or DataTable export | `InsertDataSet(...)`, `InsertDataTable(...)`, then `Save(...)` | Fast package writers can be selected automatically when the workbook shape is eligible. Check `LastSaveDiagnostics` after save to confirm the writer used or the fallback reason. |
| Object reports | `InsertObjects(...)`, table helpers, and one document-boundary save | Keep layout, AutoFit, tables, and formulas in one document session so shared strings, styles, and package finalization are batched. |
| Wide reports with AutoFit | `Execution.SaveWorksheetAfterAutoFit = false`, then `Save(...)` | Defers worksheet-part writes until the document boundary. This is the recommended report-export mode for large generated sheets. |
| Formula-backed reports | `doc.Calculate()` or `ExcelSaveOptions.EvaluateFormulasBeforeSave` | Only supported formula shapes are evaluated. Pair unsupported formulas with `ForceFullCalculationOnOpen` when the spreadsheet app should finish calculation. |

## Recommended Read Paths

| Workload | Preferred API | Notes |
| --- | --- | --- |
| Bounded range reads | `ReadRange(...)`, `Rows(...)`, or typed `RowsAs<T>(...)` | Best when callers need materialized data for a known range. |
| Very large reads | `ReadRangeStream(...)`, `RowsAsStream<T>(...)`, or `ReadObjectsStream<T>(...)` | Streams rows while keeping workbook state bounded. Prefer these when only one pass over the data is needed. |
| Unknown workbook intake | `InspectFeatures()`, `InspectFormulas()`, and targeted read options | Treat preserve-only and unsupported findings as a preflight signal before edit-heavy flows. |

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

## Measuring A Change

Use the benchmark harness for repeatable local evidence:

```powershell
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- comparison-suite --out-dir .\Docs\benchmarks\comparison-current --row-set 2500,25000 --warmup 1 --iterations 3
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- write-profile --rows 25000
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- read-profile --rows 25000
```

Use the comparison summary for public-facing numbers only when the run records:

- row counts and scenario names
- Release configuration and target framework
- raw samples, mean, median, and allocation data
- package-size and package-part metrics when save behavior matters
- machine and runtime information from the artifact manifest

## Current Boundaries

- Large workbook guidance is strongest for generated report-style workbooks and bounded read workflows.
- Feature-rich externally authored workbooks should be inspected before mutation because preserve-only package parts may need round-trip care.
- Fast package writers are automatic optimizations, not a compatibility promise for every workbook shape.
- Rendering/export is not part of the current large-workbook promise.

## Related Evidence

- Benchmark artifact guide: `Docs/benchmarks/README.md`
- Benchmark notes: `Docs/officeimo.excel.benchmark-notes.md`
- Current capability matrix: `OfficeIMO.Excel/COMPATIBILITY.md`
- Release checklist: `Docs/officeimo.excel.release-checklist.md`
