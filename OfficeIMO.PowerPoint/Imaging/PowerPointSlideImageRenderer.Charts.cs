using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static void AddChart(OfficeDrawing drawing, PowerPointChart chart, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, A.ColorScheme? colorScheme) {
            if (!TryGetBounds(chart, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
                return;
            }

            if (!chart.TryGetSnapshot(colorScheme, out PowerPointChartSnapshot snapshot)) {
                AddUnsupportedShapeDiagnostic(diagnostics, chart, "Skipped a PowerPoint chart because its cached chart data could not be converted into a shared Drawing chart snapshot.");
                return;
            }

            try {
                OfficeChartSnapshot drawingSnapshot = CreateOfficeChartSnapshot(snapshot, width, height);
                OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(drawingSnapshot, useMinimumCanvas: false);
                if (chartDrawing.Width > width || chartDrawing.Height > height) {
                    AddUnsupportedShapeDiagnostic(diagnostics, chart, "Skipped a PowerPoint chart because the shared chart renderer requires a larger drawing area than the chart frame.");
                    return;
                }

                OfficeImageFrameTransform transform = CreateChartFrameTransform(chart, left, top, width, height);
                if (transform.HasTransform) {
                    drawing.AddDrawing(chartDrawing, left, top, transform);
                } else {
                    drawing.AddDrawing(chartDrawing, left, top);
                }
            } catch (ArgumentException) {
                AddUnsupportedShapeDiagnostic(diagnostics, chart, "Skipped a PowerPoint chart because its frame is too small for safe rendering.");
            } catch (InvalidOperationException) {
                AddUnsupportedShapeDiagnostic(diagnostics, chart, "Skipped a PowerPoint chart because its cached layout could not be rendered safely.");
            }
        }

        private static OfficeImageFrameTransform CreateChartFrameTransform(PowerPointChart chart, double left, double top, double width, double height) =>
            new OfficeImageFrameTransform(
                chart.Rotation ?? 0D,
                left + (width / 2D),
                top + (height / 2D),
                chart.HorizontalFlip == true,
                chart.VerticalFlip == true);

        private static OfficeChartSnapshot CreateOfficeChartSnapshot(PowerPointChartSnapshot snapshot, double width, double height) =>
            PowerPointChartSnapshotMapper.ToOfficeSnapshot(snapshot, width, height);
    }
}
