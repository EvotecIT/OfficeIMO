using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    // Table authoring fits only its projected columns. Existing explicit widths are
    // intentional layout; auto-fit widths remain eligible for later table content.
    private void FinalizeTypedTable(ExcelTable table, OfficeTableOptions? options) {
        WriteLock(() => {
            if ((options as ExcelTableOptions)?.AutoFit != false && table.Range is { } range) {
                var (_, firstColumn, _, lastColumn) = ExcelA1.ParseRange(range);
                Column[] definitions = WorksheetRoot.GetFirstChild<Columns>()?.Elements<Column>().ToArray()
                    ?? Array.Empty<Column>();
                int[] toFit = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1)
                    .Where(index => !definitions.Any(definition => definition.Min is { } min && definition.Max is { } max
                        && min.Value <= (uint)index && max.Value >= (uint)index && definition.Width?.HasValue == true
                        && definition.BestFit?.Value != true))
                    .ToArray();
                if (toFit.Length > 0) AutoFitColumnsFor(toFit);
            }
            ApplyTableAlignments(table, options);
        });
    }

    // Explicit names also prevent deferred table registration from restoring source-member names.
    private static IReadOnlyList<string> GetTypedTableColumnNames(IReadOnlyList<string> names, bool headers, ExcelTableOptions? options) =>
        !headers && options?.UseDefaultHeaderlessColumnNames == true
            ? Enumerable.Range(1, names.Count).Select(index => "Column" + index.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToArray()
            : names;
}
