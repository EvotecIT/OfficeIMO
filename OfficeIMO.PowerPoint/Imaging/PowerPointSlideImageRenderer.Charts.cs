using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static void AddChart(OfficeDrawing drawing, PowerPointChart chart, List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping, OfficeIMO.OpenXml.Internal.OfficeXmlNode colorScheme,
            OfficeChartFidelity fidelity) {
            if (!TryGetBounds(chart, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
                return;
            }

            if (!chart.TryGetSnapshot(colorScheme, fidelity, out PowerPointChartSnapshot snapshot, out string? failureReason,
                    out IReadOnlyList<string> approximations)) {
                AddUnsupportedShapeDiagnostic(diagnostics, chart, "Skipped a PowerPoint chart because the shared chart renderer cannot draw it: " +
                    (failureReason ?? "the chart could not be projected."), diagnosticCode: OfficeDiagnosticCodes.UnsupportedChart);
                return;
            }

            try {
                OfficeChartSnapshot drawingSnapshot = CreateOfficeChartSnapshot(snapshot, width, height);
                OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(drawingSnapshot, useMinimumCanvas: false, diagnostics);
                if (chartDrawing.Width > width || chartDrawing.Height > height) {
                    AddUnsupportedShapeDiagnostic(diagnostics, chart, "Skipped a PowerPoint chart because the shared chart renderer requires a larger drawing area than the chart frame.", diagnosticCode: OfficeDiagnosticCodes.UnsupportedChart);
                    return;
                }

                if (snapshot.ProjectionNotice != null) {
                    diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                        PowerPointImageExportDiagnosticCodes.ChartProjected, snapshot.ProjectionNotice, DescribeShape(chart),
                        OfficeConversionLossKind.Approximation));
                }

                foreach (string approximation in approximations) {
                    diagnostics.Add(new OfficeImageExportDiagnostic(OfficeImageExportDiagnosticSeverity.Warning,
                        OfficeDiagnosticCodes.ChartApproximation, approximation, DescribeShape(chart),
                        OfficeConversionLossKind.Approximation));
                }

                OfficeImageFrameTransform transform = CreateChartFrameTransform(chart, left, top, width, height);
                if (transform.HasTransform) {
                    drawing.AddDrawing(chartDrawing, left, top, transform);
                } else {
                    drawing.AddDrawing(chartDrawing, left, top);
                }
            } catch (NotSupportedException error) {
                AddUnsupportedShapeDiagnostic(diagnostics, chart, "Skipped a PowerPoint chart because the shared chart renderer cannot draw it: " +
                    error.Message, diagnosticCode: OfficeDiagnosticCodes.UnsupportedChart);
            } catch (ArgumentException) {
                AddUnsupportedShapeDiagnostic(diagnostics, chart, "Skipped a PowerPoint chart because its frame is too small for safe rendering.", diagnosticCode: OfficeDiagnosticCodes.UnsupportedChart);
            } catch (InvalidOperationException) {
                AddUnsupportedShapeDiagnostic(diagnostics, chart, "Skipped a PowerPoint chart because its cached layout could not be rendered safely.", diagnosticCode: OfficeDiagnosticCodes.UnsupportedChart);
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
