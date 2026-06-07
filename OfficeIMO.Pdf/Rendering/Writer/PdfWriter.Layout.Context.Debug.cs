namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private static readonly PdfColor DebugFlowObjectBoxColor = new PdfColor(1D, 0D, 1D);
        private static readonly PdfColor DebugCanvasItemBoxColor = new PdfColor(0D, 0.65D, 1D);

        private void DrawDebugFlowObjectBox(double x, double bottomY, double boxWidth, double boxHeight) {
            if (currentOpts.Debug?.ShowFlowObjectBoxes == true) {
                DrawRowRect(sb, DebugFlowObjectBoxColor, 0.6D, x, bottomY, boxWidth, boxHeight, emitGeneratedStructure);
                pageDirty = true;
            }
        }

        private void DrawDebugCanvasItemBox(double x, double bottomY, double boxWidth, double boxHeight) {
            if (currentOpts.Debug?.ShowCanvasItemBoxes == true) {
                DrawRowRect(sb, DebugCanvasItemBoxColor, 0.6D, x, bottomY, boxWidth, boxHeight, emitGeneratedStructure);
                pageDirty = true;
            }
        }
    }
}
