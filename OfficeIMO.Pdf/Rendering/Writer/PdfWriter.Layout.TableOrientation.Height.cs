namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Cell paragraphs and oriented source marks supply their own measured line boxes.</summary>
    private static double GetTableRowInitialTextHeight(TableBlock table, int row, int columns, double leading) {
        var cells = GetTableCellLayouts(table, row, columns);
        return cells.Count > 0 && cells.All(cell => cell.Paragraphs is { Count: > 0 } ||
            cell.TextRotation != 0 && cell.OrientedRowTextHeight.HasValue)
            ? 0D : leading;
    }
}
