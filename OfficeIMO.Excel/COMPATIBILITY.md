# OfficeIMO.Excel compatibility

This matrix describes the current workbook contract. “Partial” means useful behavior exists with an explicit preservation, mutation, rendering, or interoperability boundary.

## Formats

| Format | Current contract | Boundary |
| --- | --- | --- |
| XLSX/XLSM/XLTX/XLTM | Native create, load, inspect, edit, preserve, and save through the normal `ExcelDocument` surface | Unknown or preservation-sensitive package parts are reported before edit-heavy workflows |
| XLS | First-party BIFF8 import, supported native writer subset, guarded XLS/XLSX conversion, feature reporting, package security, and password-to-open support for XOR, classic RC4, and RC4 CryptoAPI | Unsupported encryption, drawings, complete chart visuals, pivots, tables, connections, VBA/OLE, and signatures are diagnosed or blocked |
| XLSB | First-party BIFF12 import, supported native writer subset, byte-identical unchanged save, preservation-aware supported cell rewrite, and XLSB-to-XLSX conversion | Unsupported mutations and save-time transforms fail before output; XLSX bytes are never mislabeled as XLSB |

The detailed legacy contracts are documented in [XLS/XLSX compatibility](../Docs/officeimo.excel.legacy-xls-compatibility.md) and [Word/Excel interoperability](../Docs/officeimo.word-excel-interoperability.md).

## Workbook capabilities

| Area | Status | Current contract and boundary |
| --- | --- | --- |
| Create/load/save | Supported | File, byte, and caller-owned stream workflows share deterministic lifecycle behavior; remote HTTPS loading uses explicit limits and detached save semantics |
| Typed reads | Supported | Rows, streaming rows, objects, dictionaries, `DataTable`, and range projections with deterministic header matching and ambiguity diagnostics |
| Large-workbook generation and editing | Supported for documented shapes | Direct data paths, automatic eligible package writers, deferred AutoFit saves, streaming reads, file-backed Open XML editing, budgets, and cancellation are described in the [large-workbook guide](../Docs/officeimo.excel.large-workbook-guidance.md) |
| Structural edits | Supported with ownership guards | Transactional row, column, and cell shifts plus copy, move, and transpose expose bounded dry-runs, rollback, and package diagnostics; unsafe array, pivot, table, and worksheet-boundary conflicts fail before commit |
| Tables, names, filters, validation, and formatting | Broad | Table schema/resize, AutoFilter state and criteria, reference-aware names, range algebra, named styles, conditional formatting, views, print state, allowed-edit regions, ignored errors, and sparkline lifecycles have public read/mutate/remove paths |
| Charts | Broad authoring, bounded imported mutation | Common, 3-D, radar, stock, surface, combo, and dashboard charts are available; native ChartEx supports funnel, waterfall, box-and-whisker, treemap, and sunburst creation plus title, layout, placement, and OfficeIMO-authored data updates while preserving unknown imported siblings |
| Pivot tables and interactions | Broad native subset | Worksheet-source pivots cover common fields, layouts, values, grouping, filters, calculated fields, styles, shared-cache-aware source updates, refresh metadata, and native slicer/timeline views for supported PivotTable fields; unsupported imported interaction markup remains preserved |
| Query-backed tables | Guarded native subset | Native connection/query-table chains can be authored, inspected, refreshed, detached, and structurally remapped; execution requires a caller-provided host, explicit policy, bounded result materialization, cancellation, and separate opt-in for imported commands |
| Formulas | Broad authoring/inspection, bounded evaluation | One syntax tree owns A1/R1C1 conversion and reference rewriting across formula-bearing structures; inspection distinguishes authored, cached, evaluated, dirty, deferred, unsupported, and dynamic-array state; unsupported calculations remain intact for Excel |
| Templates | Broad | Typed marker binding, repeating/optional rows and sheets, Custom XML/content-control mapping, floating and native in-cell images, diagnostics, and preservation-aware relationship cloning are available; unowned imported relationship graphs remain capability-diagnosed |
| Comments and collaboration | Partial | Legacy comments and common threaded-comment read/update/resolve/remove workflows are supported; complete collaboration authoring semantics are not claimed |
| Macros and embedded payloads | Inspect/preserve/manage | VBA, package, OLE, and ActiveX payloads can be inventoried, hashed, extracted, attached/replaced/removed through bounded package operations; OfficeIMO does not execute VBA or provide a full OLE/ActiveX editor |
| Protection and encryption | Broad | OOXML password encryption plus supported legacy password import are separate from worksheet/workbook protection; permission fidelity remains format-specific |
| Feature inspection and preflight | Supported | `InspectFeatures()`, `Can`, `EnsureCan`, capability diagnostics, repair hints, and preservation reports route reads, calculation, edits, templates, rendering, and save workflows explicitly |

## Formula evaluator

`ExcelDocument.Calculate()` and `RecalculateSupportedFormulas()` evaluate the reporting-oriented subset below and write cached results. This is not a complete Excel calculation engine: unsupported formulas remain in the workbook and can be delegated to Excel through recalculation settings. Custom application functions can be registered through `ExcelCalculationOptions`.

| Family | Built-in functions |
| --- | --- |
| Aggregation | `SUM`, `AVERAGE`, `AVERAGEA`, `MIN`, `MINA`, `MAX`, `MAXA`, `COUNT`, `COUNTA`, `COUNTBLANK`, `SUBTOTAL`, `PRODUCT` |
| Conditional aggregation | `COUNTIF`, `SUMIF`, `AVERAGEIF`, `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `MINIFS`, `MAXIFS` |
| Statistics and reporting | `MEDIAN`, `LARGE`, `SMALL`, `MODE.SNGL`, `MODE`, `GEOMEAN`, `HARMEAN`, `AVEDEV`, `DEVSQ`, `SUMXMY2`, `SUMX2MY2`, `SUMX2PY2`, `SUMSQ`, `SUMPRODUCT`, `STDEV.S`, `STDEV.P`, `VAR.S`, `VAR.P`, `PERCENTILE.INC`, `PERCENTILE.EXC`, `QUARTILE.INC`, `QUARTILE.EXC`, `PERCENTRANK.INC`, `PERCENTRANK.EXC`, `RANK.EQ`, `RANK.AVG`, `COVAR`, `COVARIANCE.P`, `COVARIANCE.S`, `CORREL`, `SLOPE`, `INTERCEPT`, `RSQ`, `FORECAST.LINEAR` |
| Financial | `PMT`, `PV`, `FV`, `NPER`, `NPV` |
| Lookup and reference | `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `XMATCH`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS` |
| Mathematics | `ABS`, `SIGN`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `MROUND`, `TRUNC`, `INT`, `CEILING.MATH`, `FLOOR.MATH`, `CEILING`, `FLOOR`, `POWER`, `SQRT`, `LN`, `LOG10`, `EXP`, `PI`, `RADIANS`, `DEGREES`, `MOD` |
| Date and time | `DATE`, `TIME`, `DATEVALUE`, `TIMEVALUE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `HOUR`, `MINUTE`, `SECOND`, `DATEDIF`, `YEARFRAC`, `EDATE`, `EOMONTH`, `DAYS`, `DAYS360`, `WEEKDAY`, `WEEKNUM`, `ISOWEEKNUM`, `NETWORKDAYS`, `WORKDAY`, `WORKDAY.INTL` |
| Logical and information | `IF`, `IFS`, `SWITCH`, `CHOOSE`, `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISERROR`, `ISERR`, `ISNA`, `ISFORMULA`, `AND`, `OR`, `NOT`, `IFERROR`, `IFNA` |
| Text | `CONCAT`, `CONCATENATE`, `TEXT`, `TEXTJOIN`, `TEXTBEFORE`, `TEXTAFTER`, `FORMULATEXT`, `LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, `UPPER`, `LOWER`, `PROPER`, `SUBSTITUTE`, `FIND`, `SEARCH`, `VALUE`, `EXACT`, `REPT` |

The evaluator also handles supported arithmetic and comparison expressions, same-sheet dependencies, numeric cross-sheet references, named ranges, simple structured references, dependency-depth guards, and circular/self-reference diagnostics. Function-specific argument shapes remain bounded by focused tests; a recognized name does not imply every Excel overload, array behavior, spill behavior, or coercion rule.

## Image and PDF export

Range, worksheet, and workbook export uses shared `OfficeIMO.Drawing` primitives for PNG, JPEG, TIFF, WebP, and SVG output. The Excel adapter owns worksheet geometry, print areas, page breaks, headers/footers, cells, comments, images, charts, shapes, conditional formatting, and sparklines. Unsupported or approximate layout remains in export diagnostics; OfficeIMO does not claim to reproduce the Microsoft Excel layout engine.

PDF output uses the Excel print model and the shared PDF/Drawing owners. Route evidence and fidelity states are generated in the [PDF conversion support matrix](../Docs/officeimo.pdf-conversion-support-matrix.md).

## Security and preservation

`ExcelLoadOptions.PackageSecurity` applies shared limits and policy to Open XML, XLSB, and compound XLS inputs before parsing. Secure defaults retain compatible active content within structural limits; untrusted defaults reject active and external content. Typed findings identify the rejecting rule and part.

Unknown or preservation-sensitive content is inspected before mutation. Relationship-backed drawings, workbook structures, charts, modern-chart parts, pivot/query metadata, and template bindings are retained through load/clone/edit/save paths covered by the package-preservation corpus. A successful save is not treated as evidence that every feature became editable.

## Performance evidence

The [benchmark website](https://officeimo.com/benchmarks/) is the current cross-library evidence surface. Results remain separated by workload, format, operating system, runtime, run mode, validation contract, and source provenance. The package README does not embed a historical single-machine ranking.

## Validation

- [Excel tests](../OfficeIMO.Excel.Tests)
- [Office interoperability gate](../Build/Test-OfficeInteroperabilityGate.ps1)
- [Large-workbook guidance](../Docs/officeimo.excel.large-workbook-guidance.md)
- [Benchmark harness](../OfficeIMO.Excel.Benchmarks/README.md)
- [Image export capability matrix](../Docs/officeimo.image-export-capability-matrix.md)
