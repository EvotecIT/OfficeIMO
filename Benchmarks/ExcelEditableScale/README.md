# Editable workbooks at scale

This opt-in .NET 10 console project measures editable-workbook workloads that
ordinary read/write benchmarks leave out. It compares OfficeIMO with ClosedXML,
EPPlus, NPOI, XLibur and, where it can do the same work, the Open XML SDK. It
stays outside the normal solution, so the comparison libraries never reach
OfficeIMO's runtime dependencies.

Each process runs one library and one lane:

```powershell
dotnet build Benchmarks/ExcelEditableScale -c Release -o <output-directory>
<output-directory>/ExcelEditableScale.Benchmarks.exe <library> <lane> <iterations> [workbook.xlsx]
```

`library` is `officeimo`, `closedxml`, `epplus`, `npoi`, `xlibur` or `sdk`.
Running one pair per process keeps the peak working set specific to that
library and workload. The output line reports the median and minimum time,
allocated megabytes per operation, the process peak working set, the output
size and the validation result. The process exits with 1 when validation fails.
A pair that a library cannot run prints `not applicable`.

| Lane | Workload |
|---|---|
| `write1m` | 1,000,000 rows of number, text, decimal and Boolean cells |
| `styles` | 200,000 cells over 2,000 distinct combinations of fill, font size, bold, italic and border |
| `merged` | 50,000 rows, each with A:C merged |
| `cfdv` | 2,000 "greater than" rules and 2,000 whole-number validations over five-row blocks |
| `namestables` | 200 tables of 50 rows and 1,000 workbook-level defined names |
| `recalc` | Opens a formula workbook, changes 1,000 lookup inputs and recalculates VLOOKUP, INDEX/MATCH, IF/ROUND and SUMIFS chains |
| `autofit` | Autofit of 10 columns over 20,000 rows of text and numbers; only the autofit call is timed |
| `preserve` | Opens a workbook, edits one cell and saves; needs a workbook path |

Generate the `recalc` input with `ExcelEditableScale.Benchmarks.exe sdk recalc-fixture <path>`.
For `preserve`, use real workbooks containing charts, images and pivot caches;
the lane checks that each kind of part survives.

Validation reopens every output with the Open XML SDK, so a library cannot win
by doing less work. It checks row counts and value sums, the number of distinct
cell styles, merge counts, the cells covered by rules and validations, defined
names and table parts, positive widths for every autofit column, and part
counts before and after a preserve. `recalc` compares the summary total with
the expected value computed from the same inputs.

Libraries are used as their documentation shows for each task. NPOI creates
each style once and assigns it, and merges through `AddMergedRegionUnsafe`,
because its checked `AddMergedRegion` compares each region with every earlier
one and does not finish at this scale.
