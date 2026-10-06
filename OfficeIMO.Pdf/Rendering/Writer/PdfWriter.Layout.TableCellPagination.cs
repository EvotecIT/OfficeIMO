namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>
    /// Chooses a row fragment that fits every cell's resolved paragraph rules. Unresolved
    /// shared-table paragraphs retain the existing two-line first-fragment default.
    /// Constraints relax only when the current frame has no earlier body content to move.
    /// </summary>
    private static int LimitTableRowFragmentToParagraphBoundaries(TableCellTextLayout[] layouts,
        IReadOnlyList<TableCellLayout> cells, int startLine, int maximumLineCount, double fullFrameHeight,
        bool canMoveToNextFrame, bool requireDefaultFirstFragment = false) {
        for (int count = maximumLineCount; count > 0; count--) {
            if (IsTableRowFragmentBoundaryAllowed(layouts, cells, startLine, count, fullFrameHeight, requireDefaultFirstFragment)) return count;
        }
        return canMoveToNextFrame ? 0 : maximumLineCount;
    }

    /// <summary>Tests one boundary for both continuation fitting and first-fragment preflight.</summary>
    private static bool IsTableRowFragmentBoundaryAllowed(TableCellTextLayout[] layouts,
        IReadOnlyList<TableCellLayout> cells, int startLine, int count, double fullFrameHeight,
        bool requireDefaultFirstFragment = false) {
        int boundary = startLine + count;
        bool allowed = true;
        foreach (TableCellLayout cell in cells) {
            TableCellTextLayout layout = layouts[cell.Column];
            if (boundary >= layout.Lines.Count) continue;
            var ranges = layout.ParagraphRanges;
            if (requireDefaultFirstFragment && startLine == 0 && count < 2 &&
                (ranges == null || ranges.Any(range => !range.Paragraph.WidowControl.HasValue))) {
                allowed = false;
                break;
            }
            if (ranges == null) continue;
            for (int index = 0; index < ranges.Count; index++) {
                TableCellParagraphRange range = ranges[index];
                int end = range.StartLine + range.LineCount;
                if (boundary == end && range.Paragraph.KeepWithNext && index + 1 < ranges.Count) {
                    allowed = false;
                    break;
                }
                if (boundary <= range.StartLine || boundary >= end) continue;
                double paragraphHeight = layout.LineHeightPrefix[end] - layout.LineHeightPrefix[range.StartLine];
                if (range.Paragraph.KeepTogether && paragraphHeight <= fullFrameHeight + 0.001D) {
                    allowed = false;
                    break;
                }
                if (range.Paragraph.WidowControl == true &&
                    (boundary - Math.Max(startLine, range.StartLine) < 2 || end - boundary < 2)) {
                    allowed = false;
                    break;
                }
            }
            if (!allowed) break;
        }
        return allowed;
    }
}
