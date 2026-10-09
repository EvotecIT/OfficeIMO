using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Resolves Word's inline grid origin after effective cell margins and borders are known.</summary>
    private static void ApplyNativeInlineTablePlacement(WordTable table, TableLayout layout,
        PdfCore.PdfTableStyle style, NativeTableStyleDefaults defaults) {
        if (layout.Rows.Count == 0 || layout.Rows[0].Count == 0 || style.CellSpacing > 0D ||
            table._tableProperties?.TablePositionProperties != null) return;

        int firstColumn = GetNativeTableRowStartColumn(layout, 0);
        PdfCore.PdfAlign alignment = MapNativeTableAlignment(ResolveNativeTableAlignment(table, defaults));
        double firstMargin = GetNativeTableCellHorizontalMargin(style, 0, firstColumn, right: false);
        if (!UsesModernNativeWordLayout(table.Document)) {
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
}
