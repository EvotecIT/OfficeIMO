using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private static void FormatSimpleGenericTableSheet(ExcelSheet sheet, HtmlSemanticTable table) {
        // Generic two-column tables are often term/definition data. Keep both
        // columns within a printable width and let longer values wrap visibly.
        if (table.Rows.Count == 0 || table.Rows.Any(row => row.Cells.Count != 2
            || row.Cells.Any(cell => cell.RowSpan != 1 || cell.ColumnSpan != 1))) return;
        // The imported grid can contain empty source rows or stop at MaxTableCells.
        // Only format values that made it into the worksheet.
        ExcelCellValueInfo[] importedCells = sheet.EnumerateCells()
            .Where(cell => cell.Column is 1 or 2)
            .ToArray();
        if (importedCells.Length == 0) return;

        int firstLength = 0;
        int secondLength = 0;
        foreach (ExcelCellValueInfo cell in importedCells) {
            int length = cell.Value?.ToString()?.Length ?? 0;
            if (cell.Column == 1) firstLength = Math.Max(firstLength, length);
            else secondLength = Math.Max(secondLength, length);
        }

        double firstWidth = Math.Min(60D, Math.Max(12D, firstLength + 2D));
        double secondWidth = Math.Min(60D, Math.Max(12D, secondLength + 2D));
        const double printableWidth = 75D;
        if (firstWidth + secondWidth > printableWidth) {
            firstWidth = Math.Max(12D, Math.Ceiling(firstWidth * printableWidth / (firstWidth + secondWidth)));
            secondWidth = printableWidth - firstWidth;
        }

        sheet.SetColumnWidth(1, firstWidth);
        sheet.SetColumnWidth(2, secondWidth);
        foreach (ExcelCellValueInfo cell in importedCells) {
            sheet.CellWrapText(cell.Row, cell.Column);
        }
    }
}
