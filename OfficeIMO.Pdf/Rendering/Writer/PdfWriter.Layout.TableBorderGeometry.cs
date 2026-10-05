namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private enum TableCellBorderEdge { Top, Right, Bottom, Left }

    private static double GetCellBorderPairHalfGap(PdfCellBorderSide? side) =>
        IsRenderableCellBorderSide(side) && side!.LineStyle == PdfCellBorderLineStyle.TwoLine
            ? GetDoubleBorderGap(side.Width) / 2D : 0D;

    // Matching sides share a pair centred on the cell boundary. Reserve its
    // inward painted extent during both measurement and content placement.
    private static double GetTablePairedBorderClearance(PdfTableStyle style, int row, int column, TableCellBorderEdge edge) {
        if (style.CellBorders == null || !style.CellBorders.TryGetValue((row, column), out PdfCellBorder? border)) return 0D;
        PdfCellBorderSide? side = edge switch {
            TableCellBorderEdge.Top => border.Top ? ResolveCellBorderSide(border.TopBorderSnapshot, border) : null,
            TableCellBorderEdge.Right => border.Right ? ResolveCellBorderSide(border.RightBorderSnapshot, border) : null,
            TableCellBorderEdge.Bottom => border.Bottom ? ResolveCellBorderSide(border.BottomBorderSnapshot, border) : null,
            _ => border.Left ? ResolveCellBorderSide(border.LeftBorderSnapshot, border) : null
        };
        return IsRenderableCellBorderSide(side) && side!.LineStyle == PdfCellBorderLineStyle.TwoLine
            ? (GetDoubleBorderGap(side.Width) + side.Width) / 2D : 0D;
    }
}
