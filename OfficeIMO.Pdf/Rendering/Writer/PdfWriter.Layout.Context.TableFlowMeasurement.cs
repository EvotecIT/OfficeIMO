namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private bool measuredNonFlowTablePaint;

        private bool RestoresMeasuredTableFlow(PdfTableStyle style, double paintHeight, double availableHeight) {
            // Painting restores the cursor only when the table completes in the
            // current frame. A table that continues onto another page keeps its
            // last-page flow position, and must retain its measured advance.
            if (style.ConsumesVerticalFlow || style.Position != null || paintHeight > availableHeight + .001D)
                return false;
            measuredNonFlowTablePaint = true;
            return true;
        }
    }
}
