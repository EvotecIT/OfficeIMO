# OfficeIMO.Excel.IWork

`OfficeIMO.Excel.IWork` is the opt-in adapter for importing modern Apple Numbers files into editable `OfficeIMO.Excel` workbooks. Installing `OfficeIMO.Excel` alone does not add the iWork reader.

```bash
dotnet add package OfficeIMO.Excel.IWork
```

```csharp
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("source.numbers");
using NumbersToExcelResult result = source.ToExcelDocumentResult(
    new IWorkConversionOptions { Mode = IWorkConversionMode.Auto });

Console.WriteLine(result.Report.ProjectionKind);
Console.WriteLine(result.HasLoss);
result.Value.Save("converted.xlsx");
```

`IWorkSourceDocument.Open` reads and bounds the source independently of destination policy. `ToExcelDocument` returns the converted workbook directly; `ToExcelDocumentResult` also exposes the typed Numbers projection, diagnostics, preserved source records, and exact editable-versus-visual-fallback result. `ExcelIWorkConverter.ConvertNumbersToExcel*` provides equivalent path and stream convenience entry points.

Supported formula functions use qualified native identifiers and argument counts; see the [function subset](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md#formula-functions). Nested `OR`, `POWER`, qualified `TEXTJOIN`, `MINA`, `MAXA`, `AVERAGEA`, `OFFSET`, `PROB` and `RANDBETWEEN` retain editable expressions and typed caches. Native Numbers 14.5 exports qualify `OFFSET`, `PROB` and `RANDBETWEEN`. Local recalculation supports bounded rectangular `OFFSET` range arguments and single-cell results, finite numeric `PROB` vectors, and inclusive integer `RANDBETWEEN` bounds through ±(2^53−1); see the shared [Excel calculation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel/README.md#formula-inspection-and-calculation-policy) for limits. `MINA`, `MAXA` and `AVERAGEA` identities and argument counts are independently checked. Their saved/reopened formulas, numeric caches and local recalculation use synthetic native-format inputs; Apple-produced document, export and argument-type equivalence qualification remain open. Apple disallows direct text arguments to `MAXA`, while local Excel recalculation treats them as zero. `TEXTJOIN` uses `_xlfn.TEXTJOIN` in XLSX; source expressions retain the native function name and semantic table labels. Unsupported or ambiguous expressions retain valid caches without an editable formula; inspect `FormulaIsComplete` and `IWORK_TABLE_FORMULA_PARTIAL`. Cached values do not establish recalculation or cache freshness.

`NOT`, `TRUE`, `EXACT`, `LOWER`, `TRIM` and `UPPER` retain editable expressions with typed Boolean/text caches. Native `NOT` and `UPPER` cases qualify saved/reopened XLSX and recalculation after literal edits; source decoding independently qualifies `EXACT`, `LOWER` and `TRIM`. Synthetic `TRUE()` input demonstrates that conversion preserves a stale Boolean cache until explicit recalculation. ASCII casing and ordinary-space cleanup are covered; locale/Unicode behavior and Apple export equivalence remain unqualified.

Qualified base user-hidden rows and columns remain hidden in XLSX, while their values, formulas and styles remain editable. `result.Projection` retains the table’s `HiddenRows` and `HiddenColumns` as sorted one-based positions. Other hidden/filter/group states require the existing partial policy and retain source diagnostics; see the [visibility contract](../Docs/officeimo.iwork-support-matrix.md).

Individual row heights and column widths are carried into XLSX. Heights remain in points; widths use Excel character units. Out-of-range dimensions are rejected before destination creation.

Supported selected, region and alternating body-row solid fills become XLSX backgrounds, including styled and unstored empty cells. Selected fills override region defaults; region styling obeys the bounded destination style budget. Explicit no-fill keeps the reconstructed cell unfilled. Values and formula caches retain their types; unsupported fills and banding declarations remain diagnosed on the source projection.

Supported modern number, percentage, currency, and scientific formats keep numeric values and formula caches numeric. XLSX retains explicit decimal places, digit grouping, and minus, red, or parenthesized negative styles. Scientific cells use a single-digit mantissa and `E+00` exponent with zero-to-thirty fractional places. Nondefault scientific negative styles and grouping remain unqualified. Numbers' automatic decimal mode maps to fifteen optional fractional places (mantissa places for scientific cells) and emits `IWORK_NUMBERS_AUTOMATIC_DECIMALS_APPROXIMATED`; significant-digit selection and scientific notation can differ. Currency cells use their source three-letter identifier as a visible prefix, such as `USD -1,234.50`. Accounting negatives retain parentheses. `IWORK_NUMBERS_CURRENCY_DISPLAY_APPROXIMATED` reports that currency symbols, locale-specific placement, and accounting alignment are not reconstructed. Custom and other format families remain outside this subset. Unsupported selected formats emit `IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED` and require the partial-reconstruction policy to retain raw editable values.

Fixed abbreviated hours and minutes retain their selected duration format as `[h]"h" m"m"`. Values and formula caches stay numeric after converting source seconds to day serials; total hours and negative signs are retained. `IWORK_NUMBERS_DURATION_DISPLAY_APPROXIMATED` reports spacing and rounding limits. Fixed abbreviated day-only durations use `0"d"` over numeric day serials, including negative values. Whole days preserve their display; fractional days use numeric rounding rather than source truncation, reported by the same approximation diagnostic. Other duration styles, unit ranges (including week/day/hour) and automatic units retain source warnings and require the partial policy.

Qualified complete date/time patterns preserve editable typed dates, original formula caches and selected XLSX format codes. For example, `dd/MM/y HH:mm` becomes `dd/mm/yyyy hh:mm`; `h:mm a` becomes `h:mm am/pm`. The [source contract](../Docs/officeimo.iwork-support-matrix.md#conversion-acceptance-and-fidelity) lists all five patterns. `IWORK_NUMBERS_DATE_DISPLAY_APPROXIMATED` reports locale and calendar limits. Pre-1900 dates and values whose precision cannot survive XLSX still require visual fallback; the partial policy does not bypass these guards.

Fraction selections retain one-, two-, or three-digit denominator precision, or fixed denominators 2, 4, 8, 16, 10, and 100. XLSX uses mixed fractions with explicit minus and zero sections, such as `# ?/8;-# ?/8;0`; proper fractions have no leading zero and zero remains visible. `IWORK_NUMBERS_FRACTION_DISPLAY_APPROXIMATED` reports midpoint-rounding, equivalent-fraction normalization, and spacing differences. Values and formula caches remain numeric. Fraction precision is separate from automatic decimal formatting, so a fraction alone does not produce the automatic-decimal diagnostic. Nondefault fraction negative styles, grouping, and other controls remain unsupported.

Choose the acceptance policy explicitly when source details cannot be represented:

```csharp
var options = new IWorkConversionOptions {
    Mode = IWorkConversionMode.Auto,
    AllowPartialEditableReconstruction = true,
    RequireCompleteVisualCoverage = true
};
```

This retains bounded recoverable editable content and reports incomplete details. If editable output cannot be produced, it rejects a first-page or composite preview. Both settings default to `false`. Inspect `Report.IsPartialEditableReconstruction` before accepting the output; `Report.RequireCompleteEditableReconstruction()` rejects explicitly partial reconstruction. `Report.RequireNoLoss()` also rejects unassessed record fidelity. These policies do not bypass source limits or destination safety checks.

Table-region defaults and selected text styles preserve fonts, emphasis, foreground colors and horizontal alignment in XLSX without changing numeric values or formula caches. Unsupported paragraph layout, highlights and transparency produce `IWORK_NUMBERS_TABLE_TEXT_STYLE_PARTIAL`. Applying defaults to empty cells is bounded to 100,000 cells per table and 1,000,000 per conversion; source tables remain sparse.

Selected native top/middle/bottom cell alignment becomes XLSX vertical alignment. Four-sided cell padding remains on the source projection and produces `IWORK_NUMBERS_CELL_PADDING_OMITTED`; XLSX has no equivalent setting.

Set `NormalizeWorksheetNames = true` to use the Excel owner's rules for invalid, long, or colliding worksheet names. The result's `WorksheetMappings` links each destination name to its one-based source sheet/table position and original name. Each source table still receives its own worksheet; table-local formulas retain their references. Renames produce `IWORK_NUMBERS_WORKSHEET_RENAMED`, an approximation diagnostic. Coordinate-backed single-cell, same-target endpoint ranges and finite rectangular cross-table references bind native table identities to the actual destination worksheet names, including forward references and normalized collision suffixes. Mixed absolute and relative endpoints retain their flags. Table-local and cross-table whole-row/column and header-named axis references map to explicit ranges over the current target table body: header columns are excluded from row references, and header/footer rows are excluded from column references. `IWORK_NUMBERS_TABLE_BODY_RANGE_APPROXIMATED` reports fixed extents and lost native named labels or automatic table expansion; `RequireNoLoss()` rejects it. Other named-reference families remain unqualified. Single cells retain mixed flags and do not introduce the fixed-body approximation; endpoint ranges use one qualifier and require the same target identity. Unresolved identities or ambiguous body metadata retain caches with incomplete expressions; an empty destination body requires fallback.

```csharp
using NumbersToExcelResult normalized = source.ToExcelDocumentResult(
    new IWorkConversionOptions { NormalizeWorksheetNames = true });
foreach (NumbersWorksheetMapping mapping in normalized.WorksheetMappings) {
    Console.WriteLine($"{mapping.SourceSheetIndex}/{mapping.SourceTableIndex}: {mapping.RequestedName} -> {mapping.DestinationName}");
}
```

The path and stream convenience APIs accept cancellation after the options:

```csharp
using var cancellation = new CancellationTokenSource();
using NumbersToExcelResult cancellable = ExcelIWorkConverter.ConvertNumbersToExcelResult(
    "source.numbers", readOptions: null, conversionOptions: options,
    cancellationToken: cancellation.Token);
```

This token governs loading, projection, and destination construction. It also governs later projections from `cancellable.Source`; reopen the source with a new token after cancellation. Saving is a separate destination-owner operation.

The adapter directly depends on `OfficeIMO.Core`, `OfficeIMO.IWork`, and `OfficeIMO.Excel`. It does not add iWork support to the default Excel package graph.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for supported structures and conversion limits.

Qualified Numbers root cell comments retain plain text, display author, UTC creation time and cell address as XLSX threaded comments. Inspect `IWorkTableCell.Comment` on the source projection. Replies and unsupported comment records remain diagnosed; author normalization and empty comment text use destination fallback. See the [table-comment contract](../Docs/officeimo.iwork-support-matrix.md#table-cell-comments).

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Excel.IWork` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
