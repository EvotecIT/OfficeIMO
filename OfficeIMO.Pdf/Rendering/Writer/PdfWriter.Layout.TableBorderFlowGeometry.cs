namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Measures outward cell paint after the shared border preparation has resolved paired strokes.</summary>
    internal static (double[] Tops, double[] Bottoms) MeasureTableCellBorderFlowOutsets(TableBlock table) {
        PdfTableStyle prepared = PreparePairedTableBorders(table, table.Style!);
        int columns = GetTableColumnCount(table);
        var tops = new double[table.Rows.Count];
        var bottoms = new double[table.Rows.Count];
        for (int row = 0; row < table.Rows.Count; row++) {
            foreach (TableCellLayout cell in GetTableCellLayouts(table, row, columns)) {
                int end = Math.Min(table.Rows.Count - 1, row + cell.RowSpan - 1);
                if (prepared.CellBorders == null || !prepared.CellBorders.TryGetValue((row, cell.Column), out PdfCellBorder? border)) {
                    double defaultOutset = prepared.BorderColor.HasValue ? prepared.BorderWidth / 2D : 0D;
                    tops[row] = Math.Max(tops[row], defaultOutset);
                    bottoms[end] = Math.Max(bottoms[end], defaultOutset);
                    continue;
                }
                if (border.PaintInsideFrame) continue;
                tops[row] = Math.Max(tops[row], Outset(border.Top, border.TopBorderSnapshot, border.HiddenTopColumnSegments, cell.ColumnSpan));
                bottoms[end] = Math.Max(bottoms[end], Outset(border.Bottom, border.BottomBorderSnapshot, border.HiddenBottomColumnSegments, cell.ColumnSpan));

                double Outset(bool enabled, PdfCellBorderSide? source, System.Collections.Generic.HashSet<int>? hidden, int span) {
                    if (!enabled || hidden != null && Enumerable.Range(0, span).All(hidden.Contains)) return 0D;
                    PdfCellBorderSide? side = ResolveCellBorderSide(source, border);
                    return IsRenderableCellBorderSide(side) ? side!.Width / 2D + GetCellBorderPairOutset(side) : 0D;
                }
            }
        }
        return (tops, bottoms);
    }

    /// <summary>Omits incoming paint clearance from a neighbour on the preceding page.</summary>
    private static PdfTableStyle PrepareTableBorderContinuation(TableBlock table, PdfTableStyle prepared, int rowIndex) {
        PdfTableStyle? authored = table.Style;
        if (authored?.CellVerticalPaddingFromBorderInterior != true || authored.Position != null || !authored.ConsumesVerticalFlow ||
            GetTableRepeatHeaderRowCount(authored) > 0 || rowIndex <= 0 || rowIndex >= table.Rows.Count)
            return prepared;
        PdfTableStyle? result = null;
        foreach (TableCellLayout cell in GetTableCellLayouts(table, rowIndex, GetTableColumnCount(table))) {
            PdfCellBorder? border = null;
            prepared.CellBorders?.TryGetValue((rowIndex, cell.Column), out border);
            double top = (GetTableCellPaddingOverride(authored, rowIndex, cell.Column)?.Top ?? GetTableCellPaddingTop(authored)) +
                GetOwnTableCellVerticalBorderInset(prepared, border, true);
            if (Math.Abs(top - GetTableCellPaddingTop(prepared, rowIndex, cell.Column)) <= .001D) continue;
            result ??= prepared.Clone();
            result.CellPaddings ??= new();
            if (!result.CellPaddings.TryGetValue((rowIndex, cell.Column), out PdfCellPadding? padding))
                result.CellPaddings[(rowIndex, cell.Column)] = padding = new();
            padding.Top = top;
        }
        return result ?? prepared;
    }
}
