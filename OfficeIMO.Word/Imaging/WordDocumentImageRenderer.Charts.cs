using OfficeIMO.Drawing;

namespace OfficeIMO.Word;

internal static partial class WordDocumentImageRenderer {
    private static bool AddChart(
        WordChart chart,
        WordImageFlowContext context,
        List<OfficeImageExportDiagnostic> diagnostics) {
        if (!chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot)) {
            if (context.IsTargetPage) {
                AddDiagnostic(
                    diagnostics,
                    WordImageExportDiagnosticCodes.UnsupportedChart,
                    "Skipped a Word chart because its cached chart data could not be projected through the shared Drawing chart renderer.",
                    "Word chart");
            }
            return false;
        }

        double width = Math.Min(Math.Max(1D, snapshot.WidthPoints), context.ContentWidth);
        double height = Math.Max(1D, snapshot.HeightPoints);
        if (snapshot.WidthPoints > 0D && width < snapshot.WidthPoints) {
            height *= width / snapshot.WidthPoints;
        }
        height = Math.Min(height, context.ContentHeight);
        if (!EnsureVerticalSpace(context, height, diagnostics)) {
            return false;
        }

        if (context.IsTargetPage) {
            try {
                OfficeChartSnapshot drawingSnapshot = snapshot.WithSize(width, height);
                OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(
                    drawingSnapshot,
                    useMinimumCanvas: false, diagnostics);
                context.Drawing.AddDrawing(chartDrawing, context.Left, context.Y);
            } catch (Exception exception) when (
                exception is ArgumentException
                || exception is InvalidOperationException
                || exception is NotSupportedException
                || exception is OverflowException) {
                AddDiagnostic(
                    diagnostics,
                    WordImageExportDiagnosticCodes.UnsupportedChart,
                    "Skipped a Word chart because the shared Drawing chart renderer rejected its cached data: " + exception.Message,
                    "Word chart");
                return false;
            }
        }

        context.Y += height + ParagraphGapPoints;
        return true;
    }

}
