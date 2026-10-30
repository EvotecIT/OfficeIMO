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
            if (style.BorderWidth.ToPoints() < 0 || double.IsNaN(style.BorderWidth.ToPoints()) || double.IsInfinity(style.BorderWidth.ToPoints())) {
                throw new ArgumentException("Table border width must be a non-negative finite value.");
            }
            if (style.RowSeparatorWidth.ToPoints() < 0 || double.IsNaN(style.RowSeparatorWidth.ToPoints()) || double.IsInfinity(style.RowSeparatorWidth.ToPoints())) {
                throw new ArgumentException("Table row separator width must be a non-negative finite value.");
            }
            if (style.HeaderSeparatorWidth.ToPoints() < 0 || double.IsNaN(style.HeaderSeparatorWidth.ToPoints()) || double.IsInfinity(style.HeaderSeparatorWidth.ToPoints())) {
                throw new ArgumentException("Table header separator width must be a non-negative finite value.");
            }
            if (style.FooterSeparatorWidth.ToPoints() < 0 || double.IsNaN(style.FooterSeparatorWidth.ToPoints()) || double.IsInfinity(style.FooterSeparatorWidth.ToPoints())) {
                throw new ArgumentException("Table footer separator width must be a non-negative finite value.");
            }
            if (style.CellPaddingX.ToPoints() < 0 || double.IsNaN(style.CellPaddingX.ToPoints()) || double.IsInfinity(style.CellPaddingX.ToPoints())) {
                throw new ArgumentException("Table horizontal cell padding must be a non-negative finite value.");
            }
            if (style.CellPaddingY.ToPoints() < 0 || double.IsNaN(style.CellPaddingY.ToPoints()) || double.IsInfinity(style.CellPaddingY.ToPoints())) {
                throw new ArgumentException("Table vertical cell padding must be a non-negative finite value.");
            }
            if (style.MinRowHeight.ToPoints() < 0 || double.IsNaN(style.MinRowHeight.ToPoints()) || double.IsInfinity(style.MinRowHeight.ToPoints())) {
                throw new ArgumentException("Table minimum row height must be a non-negative finite value.");
            }
            if (style.RowMinHeights is { } rowMinHeights) {
                foreach (OfficeLength? rowMinHeight in rowMinHeights) {
                    if (rowMinHeight.HasValue && rowMinHeight.Value < OfficeLength.Zero) {
                        throw new ArgumentException("Table row minimum heights must be non-negative finite values.");
                    }
                }
            }
            if (style.SpacingBefore.ToPoints() < 0 || double.IsNaN(style.SpacingBefore.ToPoints()) || double.IsInfinity(style.SpacingBefore.ToPoints())) {
                throw new ArgumentException("Table spacing before must be a non-negative finite value.");
            }
            if (style.Caption != null && string.IsNullOrWhiteSpace(style.Caption)) {
                throw new ArgumentException("Table caption cannot be empty or whitespace.");
            }
            if (style.CaptionFontSize.HasValue && (style.CaptionFontSize.Value.ToPoints() <= 0 || double.IsNaN(style.CaptionFontSize.Value.ToPoints()) || double.IsInfinity(style.CaptionFontSize.Value.ToPoints()))) {
                throw new ArgumentException("Table caption font size must be a positive finite value.");
            }
            if (style.MinimumShrinkFontSize.HasValue && (style.MinimumShrinkFontSize.Value.ToPoints() <= 0 || double.IsNaN(style.MinimumShrinkFontSize.Value.ToPoints()) || double.IsInfinity(style.MinimumShrinkFontSize.Value.ToPoints()))) {
                throw new ArgumentException("Table minimum shrink font size must be a positive finite value.");
            }
            if (style.CaptionSpacingAfter.ToPoints() < 0 || double.IsNaN(style.CaptionSpacingAfter.ToPoints()) || double.IsInfinity(style.CaptionSpacingAfter.ToPoints())) {
                throw new ArgumentException("Table caption spacing after must be a non-negative finite value.");
            }
            if (style.SpacingAfter.ToPoints() < 0 || double.IsNaN(style.SpacingAfter.ToPoints()) || double.IsInfinity(style.SpacingAfter.ToPoints())) {
                throw new ArgumentException("Table spacing after must be a non-negative finite value.");
            }
            if (style.PageContinuationSpacingBefore.ToPoints() < 0 || double.IsNaN(style.PageContinuationSpacingBefore.ToPoints()) || double.IsInfinity(style.PageContinuationSpacingBefore.ToPoints())) {
                throw new ArgumentException("Table page continuation spacing before must be a non-negative finite value.");
            }
            if (double.IsNaN(style.RowBaselineOffset.ToPoints()) || double.IsInfinity(style.RowBaselineOffset.ToPoints())) {
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
                    if (cellBorder.Value == null || cellBorder.Value.Width.ToPoints() < 0 || double.IsNaN(cellBorder.Value.Width.ToPoints()) || double.IsInfinity(cellBorder.Value.Width.ToPoints())) {
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
