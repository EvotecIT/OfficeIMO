namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Oriented source mark metrics supply their own row minimum; ordinary cells retain the default line box.</summary>
    private static double GetTableRowInitialTextHeight(TableBlock table, int row, int columns, double leading) {
        var cells = GetTableCellLayouts(table, row, columns);
        return cells.Count > 0 && cells.All(cell => cell.TextRotation != 0 && cell.OrientedRowTextHeight.HasValue)
            ? 0D : leading;
    }
}
