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
| Forward-only CSV or Excel reads | `TabularReader.Open(...)` | One `DbDataReader` contract covers CSV, TSV, XLSX, XLSM, and XLSB. Workbook used ranges are discovered automatically. |
| Typed object reads | `TabularReader.ReadRecords<T>()` | Column names match writable properties case-insensitively. Use `[DataMember(Name = "...")]` for headers that are not CLR property names. |
| Multiple worksheets | `TableNames`, `TableName`, and `NextResult()` | Results stay in workbook order. Set `TabularReadOptions.TableName` when only one worksheet should be opened. |
| Unknown workbook edit intake | `ExcelDocument.Load(...)`, `InspectFeatures()`, and `InspectFormulas()` | Use the editable document model only when the workbook will be inspected, mutated, converted, or saved again. Treat preserve-only and unsupported findings as a preflight signal. |

Example:

```csharp
using OfficeIMO.Tabular;

using var reader = TabularReader.Open("sales.xlsx");
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

The pinned tabular suite reproduces the public 65K-record CSV, XLSX, and XLSB
workloads with hash-verified inputs and equivalent typed getters:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Tabular.Benchmarks -- --validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Tabular.Benchmarks -- --artifacts .\artifacts\tabular-full
```

Windows, Linux, and macOS are separate evidence lanes. Do not average them or
substitute one platform when another platform is missing.

## Current Boundaries

- Large workbook guidance is strongest for generated report-style workbooks and bounded read workflows.
- Feature-rich externally authored workbooks should be inspected before mutation because preserve-only package parts may need round-trip care.
- Fast package writers are automatic optimizations, not a compatibility promise for every workbook shape.
- Rendering/export is not part of the current large-workbook promise.

## Related Evidence

- Benchmark artifact guide: `Docs/benchmarks/README.md`
- Benchmark notes: `Docs/officeimo.excel.benchmark-notes.md`
- Current capability matrix: `OfficeIMO.Excel/COMPATIBILITY.md`
