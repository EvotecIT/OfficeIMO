namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private enum TableCellBorderEdge { Top, Right, Bottom, Left }

    private static double GetCellBorderPairOutset(PdfCellBorderSide? side) =>
        IsRenderableCellBorderSide(side) && side!.LineStyle == PdfCellBorderLineStyle.TwoLine
            ? side.CenteredPair ? GetDoubleBorderGap(side.Width) / 2D : 0D : 0D;

    private static double GetCellBorderPairInset(PdfCellBorderSide? side) =>
        IsRenderableCellBorderSide(side) && side!.LineStyle == PdfCellBorderLineStyle.TwoLine
            ? GetDoubleBorderGap(side.Width) - GetCellBorderPairOutset(side) : 0D;

    private static bool HasPairedCellBorder(PdfCellBorder border) =>
        HasPair(border.Top, ResolveCellBorderSide(border.TopBorderSnapshot, border)) ||
        HasPair(border.Right, ResolveCellBorderSide(border.RightBorderSnapshot, border)) ||
        HasPair(border.Bottom, ResolveCellBorderSide(border.BottomBorderSnapshot, border)) ||
        HasPair(border.Left, ResolveCellBorderSide(border.LeftBorderSnapshot, border));

    private static bool HasPair(bool enabled, PdfCellBorderSide? side) =>
        enabled && IsRenderableCellBorderSide(side) && side!.LineStyle == PdfCellBorderLineStyle.TwoLine;

    // The prepared snapshot reserves both authored and incoming paired paint.
    private static double GetTablePairedBorderClearance(PdfTableStyle style, int row, int column, TableCellBorderEdge edge) {
        if (style.CellBorders == null || !style.CellBorders.TryGetValue((row, column), out PdfCellBorder? border)) return 0D;
        PdfCellBorderSide? side = edge switch {
            TableCellBorderEdge.Top => border.Top ? ResolveCellBorderSide(border.TopBorderSnapshot, border) : null,
            TableCellBorderEdge.Right => border.Right ? ResolveCellBorderSide(border.RightBorderSnapshot, border) : null,
            TableCellBorderEdge.Bottom => border.Bottom ? ResolveCellBorderSide(border.BottomBorderSnapshot, border) : null,
            _ => border.Left ? ResolveCellBorderSide(border.LeftBorderSnapshot, border) : null
        };
        return IsRenderableCellBorderSide(side) && side!.LineStyle == PdfCellBorderLineStyle.TwoLine
            ? GetCellBorderPairInset(side) + side.Width / 2D : 0D;
    }

    private static PdfTableStyle PreparePairedTableBorders(TableBlock table, PdfTableStyle style) {
        if (style.CellBorders == null || !style.CellBorders.Values.Any(HasPairedCellBorder)) return style;

        PdfTableStyle prepared = style.Clone();
        int columns = GetTableColumnCount(table);
        var owners = new System.Collections.Generic.Dictionary<(int Row, int Column), (int Row, int Column)>();
        var layouts = new System.Collections.Generic.Dictionary<(int Row, int Column), TableCellLayout>();
        for (int row = 0; row < table.Rows.Count; row++) {
            foreach (TableCellLayout cell in GetTableCellLayouts(table, row, columns)) {
                var anchor = (row, cell.Column);
                layouts[anchor] = cell;
                for (int r = row; r < row + cell.RowSpan; r++)
                    for (int c = cell.Column; c < cell.Column + cell.ColumnSpan; c++) owners[(r, c)] = anchor;
            }
        }
        var borders = prepared.CellBorders!;
        foreach (var entry in borders) {
            if (!layouts.TryGetValue(entry.Key, out TableCellLayout cell)) continue;
            PdfCellBorder border = entry.Value;
            border.TopBorder = PrepareSide(border.Top, border.TopBorderSnapshot, border, entry.Key.Row > 0);
            border.RightBorder = PrepareSide(border.Right, border.RightBorderSnapshot, border, cell.Column + cell.ColumnSpan < columns);
            border.BottomBorder = PrepareSide(border.Bottom, border.BottomBorderSnapshot, border, entry.Key.Row + cell.RowSpan < table.Rows.Count);
            border.LeftBorder = PrepareSide(border.Left, border.LeftBorderSnapshot, border, cell.Column > 0);
        }
        double spacing = GetTableCellSpacing(prepared);
        prepared.CellPaddings ??= new System.Collections.Generic.Dictionary<(int Row, int Column), PdfCellPadding>();
        foreach (var entry in borders) {
            if (!layouts.TryGetValue(entry.Key, out TableCellLayout cell)) continue;
            PdfCellBorder border = entry.Value;
            for (int segment = 0; segment < cell.ColumnSpan; segment++) {
                ReserveNeighbour(border.Top, border.TopBorderSnapshot, border.HiddenTopColumnSegments, segment,
                    (entry.Key.Row - 1, cell.Column + segment), TableCellBorderEdge.Bottom);
                ReserveNeighbour(border.Bottom, border.BottomBorderSnapshot, border.HiddenBottomColumnSegments, segment,
                    (entry.Key.Row + cell.RowSpan, cell.Column + segment), TableCellBorderEdge.Top);
            }
            for (int segment = 0; segment < cell.RowSpan; segment++) {
                ReserveNeighbour(border.Left, border.LeftBorderSnapshot, border.HiddenLeftRowSegments, segment,
                    (entry.Key.Row + segment, cell.Column - 1), TableCellBorderEdge.Right);
                ReserveNeighbour(border.Right, border.RightBorderSnapshot, border.HiddenRightRowSegments, segment,
                    (entry.Key.Row + segment, cell.Column + cell.ColumnSpan), TableCellBorderEdge.Left);
            }
        }
        return prepared;

        static PdfCellBorderSide? PrepareSide(bool enabled, PdfCellBorderSide? side, PdfCellBorder border, bool internalBoundary) {
            PdfCellBorderSide? resolved = ResolveCellBorderSide(side, border);
            if (!HasPair(enabled, resolved)) return side;
            var result = resolved!.Clone();
            result.CenteredPair = internalBoundary;
            return result;
        }
        void ReserveNeighbour(bool enabled, PdfCellBorderSide? side, System.Collections.Generic.HashSet<int>? hidden, int segment,
            (int Row, int Column) coordinate, TableCellBorderEdge edge) {
            if (!HasPair(enabled, side) || hidden?.Contains(segment) == true || !owners.TryGetValue(coordinate, out var anchor)) return;
            double clearance = Math.Max(0D, GetCellBorderPairOutset(side) + side!.Width / 2D - spacing);
            if (clearance <= 0D) return;
            if (!prepared.CellPaddings.TryGetValue(anchor, out PdfCellPadding? padding)) {
                padding = new PdfCellPadding();
                prepared.CellPaddings[anchor] = padding;
            }
            switch (edge) {
                case TableCellBorderEdge.Top: padding.Top = Math.Max(GetTableCellPaddingTop(prepared, anchor.Row, anchor.Column), clearance); break;
                case TableCellBorderEdge.Right: padding.Right = Math.Max(GetTableCellPaddingRight(prepared, anchor.Row, anchor.Column), clearance); break;
                case TableCellBorderEdge.Bottom: padding.Bottom = Math.Max(GetTableCellPaddingBottom(prepared, anchor.Row, anchor.Column), clearance); break;
                default: padding.Left = Math.Max(GetTableCellPaddingLeft(prepared, anchor.Row, anchor.Column), clearance); break;
            }
        }
    }
}
