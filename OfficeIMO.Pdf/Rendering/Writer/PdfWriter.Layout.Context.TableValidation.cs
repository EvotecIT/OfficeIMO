namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Validates the table contract before preparing geometry for either body or note flow.</summary>
        private static void ValidateFlowTable(TableBlock tb, PdfTableStyle style, int cols) {
            if (!IsValidPdfAlign(tb.Align)) {
                throw new ArgumentException("Table alignment must be Left, Center, or Right.");
            }
            if (style.Alignments != null) {
                foreach (var alignment in style.Alignments) {
                    if (!IsValidPdfColumnAlign(alignment)) {
                        throw new ArgumentException("Table column alignments must be Left, Center, or Right.");
                    }
                }
            }
            if (style.VerticalAlignments != null) {
                foreach (var verticalAlignment in style.VerticalAlignments) {
                    if (!IsValidPdfCellVerticalAlign(verticalAlignment)) {
                        throw new ArgumentException("Table vertical alignments must be defined PDF cell vertical alignment values.");
                    }
                }
            }
            if (!IsValidPdfAlign(style.CaptionAlign)) {
                throw new ArgumentException("Table caption alignment must be Left, Center, or Right.");
            }
            if (style.BorderWidth < 0 || double.IsNaN(style.BorderWidth) || double.IsInfinity(style.BorderWidth)) {
                throw new ArgumentException("Table border width must be a non-negative finite value.");
            }
            if (style.RowSeparatorWidth < 0 || double.IsNaN(style.RowSeparatorWidth) || double.IsInfinity(style.RowSeparatorWidth)) {
                throw new ArgumentException("Table row separator width must be a non-negative finite value.");
            }
            if (style.HeaderSeparatorWidth < 0 || double.IsNaN(style.HeaderSeparatorWidth) || double.IsInfinity(style.HeaderSeparatorWidth)) {
                throw new ArgumentException("Table header separator width must be a non-negative finite value.");
            }
            if (style.FooterSeparatorWidth < 0 || double.IsNaN(style.FooterSeparatorWidth) || double.IsInfinity(style.FooterSeparatorWidth)) {
                throw new ArgumentException("Table footer separator width must be a non-negative finite value.");
            }
            if (style.CellPaddingX < 0 || double.IsNaN(style.CellPaddingX) || double.IsInfinity(style.CellPaddingX)) {
                throw new ArgumentException("Table horizontal cell padding must be a non-negative finite value.");
            }
            if (style.CellPaddingY < 0 || double.IsNaN(style.CellPaddingY) || double.IsInfinity(style.CellPaddingY)) {
                throw new ArgumentException("Table vertical cell padding must be a non-negative finite value.");
            }
            if (style.MinRowHeight < 0 || double.IsNaN(style.MinRowHeight) || double.IsInfinity(style.MinRowHeight)) {
                throw new ArgumentException("Table minimum row height must be a non-negative finite value.");
            }
            if (style.RowMinHeights != null) {
                foreach (double? rowMinHeight in style.RowMinHeights) {
                    if (rowMinHeight.HasValue && (rowMinHeight.Value < 0 || double.IsNaN(rowMinHeight.Value) || double.IsInfinity(rowMinHeight.Value))) {
                        throw new ArgumentException("Table row minimum heights must be non-negative finite values.");
                    }
                }
            }
            if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
                throw new ArgumentException("Table spacing before must be a non-negative finite value.");
            }
            if (style.Caption != null && string.IsNullOrWhiteSpace(style.Caption)) {
                throw new ArgumentException("Table caption cannot be empty or whitespace.");
            }
            if (style.CaptionFontSize.HasValue && (style.CaptionFontSize.Value <= 0 || double.IsNaN(style.CaptionFontSize.Value) || double.IsInfinity(style.CaptionFontSize.Value))) {
                throw new ArgumentException("Table caption font size must be a positive finite value.");
            }
            if (style.MinimumShrinkFontSize.HasValue && (style.MinimumShrinkFontSize.Value <= 0 || double.IsNaN(style.MinimumShrinkFontSize.Value) || double.IsInfinity(style.MinimumShrinkFontSize.Value))) {
                throw new ArgumentException("Table minimum shrink font size must be a positive finite value.");
            }
            if (style.CaptionSpacingAfter < 0 || double.IsNaN(style.CaptionSpacingAfter) || double.IsInfinity(style.CaptionSpacingAfter)) {
                throw new ArgumentException("Table caption spacing after must be a non-negative finite value.");
            }
            if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
                throw new ArgumentException("Table spacing after must be a non-negative finite value.");
            }
            if (style.PageContinuationSpacingBefore < 0 || double.IsNaN(style.PageContinuationSpacingBefore) || double.IsInfinity(style.PageContinuationSpacingBefore)) {
                throw new ArgumentException("Table page continuation spacing before must be a non-negative finite value.");
            }
            if (double.IsNaN(style.RowBaselineOffset) || double.IsInfinity(style.RowBaselineOffset)) {
                throw new ArgumentException("Table row baseline offset must be a finite value.");
            }
            if (style.CellFills != null) {
                foreach (var cellFill in style.CellFills) {
                    if (cellFill.Key.Row < 0 || cellFill.Key.Column < 0) {
                        throw new ArgumentException("Table cell fill coordinates cannot be negative.");
                    }
                }
            }
            if (style.CellBorders != null) {
                foreach (var cellBorder in style.CellBorders) {
                    if (cellBorder.Key.Row < 0 || cellBorder.Key.Column < 0) {
                        throw new ArgumentException("Table cell border coordinates cannot be negative.");
                    }
                    if (cellBorder.Value == null || cellBorder.Value.Width < 0 || double.IsNaN(cellBorder.Value.Width) || double.IsInfinity(cellBorder.Value.Width)) {
                        throw new ArgumentException("Table cell border widths must be non-negative finite values.");
                    }
                }
            }
            if (style.HeaderRowCount < 0) {
                throw new ArgumentException("Table header row count cannot be negative.");
            }
            if (style.FooterRowCount < 0) {
                throw new ArgumentException("Table footer row count cannot be negative.");
            }

            ValidateTableRoleRowCounts(style, tb.Rows.Count);
            int headerRowCount = style.HeaderRowCount;
            int footerStartRowIndex = tb.Rows.Count - style.FooterRowCount;
            ValidateTableCellStyleCoordinates(style, tb, cols);
            ValidateTableColumnStyleBounds(style, cols);
            ValidateTableRowStyleBounds(style, tb.Rows.Count);
            ValidateTableRowSpansWithinRoleBoundaries(tb, cols, headerRowCount, footerStartRowIndex);

        }
    }
}
