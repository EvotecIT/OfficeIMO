namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Measures the physical width occupied by the logical lines of a turned cell.</summary>
    private static double MeasureOrientedTableTextWidth(TableCellLayout cell, PdfStandardFont font,
        double size, double leading, PdfOptions? options) {
        // Horizontal text advance becomes vertical after the turn. Automatic
        // grids reserve cross-axis line boxes, including authored spacing,
        // rather than expanding the column for the entire horizontal string.
        TableCellTextLayout layout = CreateTableCellTextLayout(cell, TableCellNoWrapWidth,
            font, size, leading, options);
        return MeasureTableCellTextHeight(layout, 0, layout.LineCount, leading);
    }
}
