namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Measures the physical width occupied by the first logical line of a turned cell.</summary>
    private static double MeasureOrientedTableTextWidth(TableCellLayout cell, PdfStandardFont font,
        double size, double leading, PdfOptions? options) {
        // Horizontal text advance becomes vertical after the turn. Automatic
        // grids reserve one cross-axis line box, including authored spacing,
        // rather than expanding the column for the entire horizontal string.
        TableCellTextLayout layout = CreateTableCellTextLayout(cell, TableCellNoWrapWidth,
            font, size, leading, options);
        return MeasureTableCellTextHeight(layout, 0, 1, leading);
    }
}
