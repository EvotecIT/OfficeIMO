namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Measures a fresh row-column frame with the same reservations and consumed top padding as group resumption.</summary>
        private (double Height, double ConsumedBefore) GetRowColumnContinuationFrame(IReadOnlyList<ColumnGroup> groups) {
            var padding = MeasureColumnGroupContinuationPadding(groups);
            return (Math.Max(0D, GetMeasuredContinuationFrameHeight() - padding.Total), padding.Top);
        }

        // A row completes on its first visual when each column holds one visual
        // whose first and full measurements agree. Later content keeps its capacity.
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
                // Both probes include leading spacing, which a column start would
                // otherwise suppress and could coincidentally equal a later line.
                if (Math.Abs(MeasureColItemFullHeight(visual, consumedBefore: 1D) - MeasureColItemFirstVisualHeight(visual)) > .001D) return false;
            }
            return true;
        }
    }
}
