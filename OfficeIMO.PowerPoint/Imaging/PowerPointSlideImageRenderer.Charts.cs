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

        private static OfficeChartSnapshot CreateOfficeChartSnapshot(PowerPointChartSnapshot snapshot, double width, double height) {
            var series = new List<OfficeChartSeries>(snapshot.Data.Series.Count);
            for (int i = 0; i < snapshot.Data.Series.Count; i++) {
                PowerPointChartSeries item = snapshot.Data.Series[i];
                series.Add(new OfficeChartSeries(
                    item.Name,
                    item.Values,
                    item.XValues,
                    color: item.Color,
                    pointColors: null,
                    showMarkers: true,
                    strokeWidth: item.StrokeWidth,
                    renderKind: MapChartKind(item.ChartKind ?? snapshot.ChartKind),
                    axisGroup: item.AxisGroup));
            }

            return new OfficeChartSnapshot(
                snapshot.Name,
                snapshot.Title,
                MapChartKind(snapshot.ChartKind),
                new OfficeChartData(snapshot.Data.Categories, series),
                width,
                height);
        }

        private static OfficeChartKind MapChartKind(PowerPointChartSnapshotKind kind) {
            switch (kind) {
                case PowerPointChartSnapshotKind.ClusteredColumn:
                    return OfficeChartKind.ColumnClustered;
                case PowerPointChartSnapshotKind.StackedColumn:
                    return OfficeChartKind.ColumnStacked;
                case PowerPointChartSnapshotKind.StackedColumn100:
                    return OfficeChartKind.ColumnStacked100;
                case PowerPointChartSnapshotKind.ClusteredBar:
                    return OfficeChartKind.BarClustered;
                case PowerPointChartSnapshotKind.StackedBar:
                    return OfficeChartKind.BarStacked;
                case PowerPointChartSnapshotKind.StackedBar100:
                    return OfficeChartKind.BarStacked100;
                case PowerPointChartSnapshotKind.Line:
                    return OfficeChartKind.Line;
                case PowerPointChartSnapshotKind.StackedLine:
                    return OfficeChartKind.LineStacked;
                case PowerPointChartSnapshotKind.StackedLine100:
                    return OfficeChartKind.LineStacked100;
                case PowerPointChartSnapshotKind.Scatter:
                    return OfficeChartKind.Scatter;
                case PowerPointChartSnapshotKind.Pie:
                    return OfficeChartKind.Pie;
                case PowerPointChartSnapshotKind.Doughnut:
                    return OfficeChartKind.Doughnut;
                case PowerPointChartSnapshotKind.Area:
                    return OfficeChartKind.Area;
                case PowerPointChartSnapshotKind.StackedArea:
                    return OfficeChartKind.AreaStacked;
                case PowerPointChartSnapshotKind.StackedArea100:
                    return OfficeChartKind.AreaStacked100;
                case PowerPointChartSnapshotKind.Radar:
                    return OfficeChartKind.Radar;
                default:
                    throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported PowerPoint chart snapshot kind.");
            }
        }
    }
}
