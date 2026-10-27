using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Resolves Word's inline grid origin after effective cell margins and borders are known.</summary>
    private static void ApplyNativeInlineTablePlacement(WordTable table, TableLayout layout,
        PdfCore.PdfTableStyle style, NativeTableStyleDefaults defaults, bool containedTable = false) {
        if (layout.Rows.Count == 0 || layout.Rows[0].Count == 0 || style.CellSpacing > 0D ||
            style.BorderFrame?.Border != null ||
            table._tableProperties?.TablePositionProperties != null) return;

        // Body legacy placement translates its text margin into a page grid.
        // A contained table already receives its physical cell content frame.
        if (!containedTable) ApplyNativeLegacyPercentageTableDefaultIndent(table, layout, style, defaults);
        int firstColumn = GetNativeTableRowStartColumn(layout, 0);
        PdfCore.PdfAlign alignment = MapNativeTableAlignment(ResolveNativeTableAlignment(table, defaults));
        double firstMargin = GetNativeTableCellHorizontalMargin(style, 0, firstColumn, right: false);
        if (!containedTable && !UsesModernNativeWordLayout(table.Document)) {
            // Legacy left/right alignment uses the leading margin of the first
            // cell as a grid translation; the authored grid widths stay intact.
            if (alignment == PdfCore.PdfAlign.Right) style.HorizontalOffset += firstMargin;
            else if (alignment != PdfCore.PdfAlign.Center) style.HorizontalOffset -= Math.Max(firstMargin,
                GetNativeTableCellBorderInset(style, 0, firstColumn, right: false));
        } else if (alignment == PdfCore.PdfAlign.Right) {
            int lastColumn = firstColumn;
            int column = firstColumn;
            foreach (WordTableCell cell in layout.Rows[0]) {
                if (IsNativeHorizontalMergeContinuation(cell)) continue;
                lastColumn = column;
                column += GetNativeCellColumnSpan(cell);
            }
            style.HorizontalOffset -= GetNativeTableCellBorderInset(style, 0, lastColumn, right: true);
        } else if (alignment != PdfCore.PdfAlign.Center) {
            style.HorizontalOffset += GetNativeTableCellBorderInset(style, 0, firstColumn, right: false);
        }

        // Cell layout keeps zero/small margins inside the border paint.
        // Preserve larger authored margins rather than adding border width to them.
        for (int row = 0; row < layout.Rows.Count; row++) {
            int column = GetNativeTableRowStartColumn(layout, row);
            foreach (WordTableCell cell in layout.Rows[row]) {
                if (IsNativeHorizontalMergeContinuation(cell)) continue;
                if (!IsNativeVerticalMergeContinuation(cell)) {
                    double left = GetNativeTableCellHorizontalMargin(style, row, column, right: false);
                    double right = GetNativeTableCellHorizontalMargin(style, row, column, right: true);
                    double leftInset = GetNativeTableCellBorderInset(style, row, column, right: false);
                    double rightInset = GetNativeTableCellBorderInset(style, row, column, right: true);
                    if (left < leftInset || right < rightInset) {
                        style.CellPaddings ??= new();
                        if (!style.CellPaddings.TryGetValue((row, column), out PdfCore.PdfCellPadding? padding))
                            style.CellPaddings[(row, column)] = padding = new();
                        padding.Left = Math.Max(left, leftInset);
                        padding.Right = Math.Max(right, rightInset);
                    }
                }
                column += GetNativeCellColumnSpan(cell);
            }
        }
    }

    private static double GetNativeTableCellHorizontalMargin(PdfCore.PdfTableStyle style, int row, int column, bool right) {
        PdfCore.PdfCellPadding? padding = null;
        style.CellPaddings?.TryGetValue((row, column), out padding);
        return right ? padding?.Right ?? style.CellPaddingRight ?? style.CellPaddingX
            : padding?.Left ?? style.CellPaddingLeft ?? style.CellPaddingX;
    }

    private static double GetNativeTableCellBorderInset(PdfCore.PdfTableStyle style, int row, int column, bool right) {
        if (style.CellBorders?.TryGetValue((row, column), out PdfCore.PdfCellBorder? border) == true)
            return (GetNativeFrameSide(border, top: null, right: right)?.PaintThickness ?? 0D) / 2D;
        return style.BorderColor.HasValue ? style.BorderWidth / 2D : 0D;
    }

    // Word never lays a column out narrower than its cell margins: the margins
    // stay and the column grows to keep a minimal text area. A column must keep
    // at least this much text width beside its left and right margins.
    private const double NativeMinimumCellTextWidth = 1D;

    private static void ApplyNativeCellMarginMinimumWidths(TableLayout layout, PdfCore.PdfTableStyle style) {
        int columns = layout.ColumnWidths.Length;
        if (columns == 0) return;
        var required = new double[columns];
        foreach (var cell in EnumerateNativeTableCellPositions(layout)) {
            if (cell.ColumnSpan != 1 || cell.Column < 0 || cell.Column >= columns) continue;
            double margins = GetNativeTableCellHorizontalMargin(style, cell.Row, cell.Column, right: false) +
                GetNativeTableCellHorizontalMargin(style, cell.Row, cell.Column, right: true);
            required[cell.Column] = Math.Max(required[cell.Column], margins + NativeMinimumCellTextWidth);
        }
        for (int column = 0; column < columns; column++) {
            if (required[column] <= 0D) continue;
            style.ColumnMinWidthPoints ??= new List<double?>();
            while (style.ColumnMinWidthPoints.Count < columns) style.ColumnMinWidthPoints.Add(null);
            if ((style.ColumnMinWidthPoints[column] ?? 0D) < required[column])
                style.ColumnMinWidthPoints[column] = required[column];
            if (style.ColumnMaxWidthPoints != null && column < style.ColumnMaxWidthPoints.Count &&
                style.ColumnMaxWidthPoints[column] is double maximum && maximum < required[column])
                style.ColumnMaxWidthPoints[column] = required[column];
            if (style.ColumnWidthPoints != null && column < style.ColumnWidthPoints.Count &&
                style.ColumnWidthPoints[column] is double width && width < required[column])
                style.ColumnWidthPoints[column] = required[column];
        }
        if (style.MaxWidth is double maxWidth) {
            double minimumTotal = style.ColumnMinWidthPoints?.Sum(width => width ?? 0D) ?? 0D;
            double gaps = Math.Max(0, columns - 1) * style.CellSpacing;
            if (maxWidth < minimumTotal + gaps) style.MaxWidth = minimumTotal + gaps;
        }
    }
}
