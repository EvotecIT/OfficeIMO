namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Measures a fresh row-column frame with the same reservations and consumed top padding as group resumption.</summary>
        private (double Height, double ConsumedBefore) GetRowColumnContinuationFrame(IReadOnlyList<ColumnGroup> groups) {
            var padding = MeasureColumnGroupContinuationPadding(groups);
            return (Math.Max(0D, GetMeasuredContinuationFrameHeight() - padding.Total), padding.Top);
        }

        // Completion is safe when each column has one visual whose prepared first
        // and full measurements agree. Later content keeps its fragment capacity.
        private bool CanRowCompleteOnItsFirstVisual(IReadOnlyList<List<ColItem>> columns) {
            foreach (var items in columns) {
                ColItem? visual = null;
                foreach (ColItem item in items) {
                    if (item is ColBookmark) continue;
                    if (item is ColGroupStart start && start.Group.Style == null) continue;
                    if (item is ColGroupEnd end && end.Group.Style == null) continue;
                    if (visual != null) return false;
                    visual = item;
                }
                if (visual == null) continue;
                // Both probes include leading spacing. Otherwise spacing suppressed
                // at column start could accidentally equal a later line's advance.
                double fullHeight = MeasureColItemFullHeight(visual, consumedBefore: 1D);
                if (Math.Abs(fullHeight - MeasureColItemFirstVisualHeight(visual)) > .001D) return false;
            }
            return true;
        }
    }
}
