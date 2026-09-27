using System;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>Owns native snapshot projection for public, image, and PDF chart consumers.</summary>
    internal static class PowerPointChartSnapshotMapper {
        internal static OfficeChartSnapshot ToOfficeSnapshot(PowerPointChartSnapshot snapshot,
            double width, double height, OfficeChartStyle? style = null, OfficeChartLayout? layout = null) {
            var series = snapshot.Data.Series.Select(item => item.BubbleSizes != null
                ? OfficeChartSeries.CreateBubble(item.Name, item.XValues!, item.Values,
                    item.BubbleSizes, item.Color, item.PointColors,
                    showInLegend: item.ShowInLegend,
                    markerOutlineColor: item.StrokeColor ?? item.Color,
                    markerOutlineWidth: item.StrokeWidth,
                    showMarkerOutline: item.ShowStroke)
                : new OfficeChartSeries(item.Name, item.Values, item.XValues, item.Color,
                    pointColors: item.PointColors, showMarkers: item.SharedAppearance?.ShowMarkers ?? true,
                    showInLegend: item.ShowInLegend, connectLine: item.SharedAppearance?.ConnectLine ?? true,
                    markerSize: item.SharedAppearance?.MarkerSize,
                    markerShape: item.SharedAppearance?.MarkerShape,
                    markerOutlineColor: item.SharedAppearance?.MarkerOutlineColor,
                    markerOutlineWidth: item.SharedAppearance?.MarkerOutlineWidth,
                    strokeWidth: item.StrokeWidth,
                    strokeDashStyle: item.SharedAppearance?.StrokeDashStyle,
                    renderKind: item.ChartKind.HasValue ? MapKind(item.ChartKind.Value) : null,
                    axisGroup: item.AxisGroup)).Select((series, index) =>
                        series.WithPointStyles(snapshot.Data.Series[index].PointStyles)).ToList();
            return new OfficeChartSnapshot(snapshot.Name, snapshot.Title, MapKind(snapshot.ChartKind),
                new OfficeChartData(snapshot.Data.Categories, series), width, height,
                style ?? snapshot.Style, layout ?? snapshot.Layout,
                bubbleScalePercent: snapshot.BubbleScalePercent, bubbleSizeMode: snapshot.BubbleSizeMode,
                radialLayout: snapshot.RadialLayout);
        }

        internal static OfficeChartKind MapKind(PowerPointChartSnapshotKind kind) => kind switch {
            PowerPointChartSnapshotKind.ClusteredColumn => OfficeChartKind.ColumnClustered,
            PowerPointChartSnapshotKind.StackedColumn => OfficeChartKind.ColumnStacked,
            PowerPointChartSnapshotKind.StackedColumn100 => OfficeChartKind.ColumnStacked100,
            PowerPointChartSnapshotKind.ClusteredBar => OfficeChartKind.BarClustered,
            PowerPointChartSnapshotKind.StackedBar => OfficeChartKind.BarStacked,
            PowerPointChartSnapshotKind.StackedBar100 => OfficeChartKind.BarStacked100,
            PowerPointChartSnapshotKind.Line => OfficeChartKind.Line,
            PowerPointChartSnapshotKind.StackedLine => OfficeChartKind.LineStacked,
            PowerPointChartSnapshotKind.StackedLine100 => OfficeChartKind.LineStacked100,
            PowerPointChartSnapshotKind.Area => OfficeChartKind.Area,
            PowerPointChartSnapshotKind.StackedArea => OfficeChartKind.AreaStacked,
            PowerPointChartSnapshotKind.StackedArea100 => OfficeChartKind.AreaStacked100,
            PowerPointChartSnapshotKind.Scatter => OfficeChartKind.Scatter,
            PowerPointChartSnapshotKind.Radar => OfficeChartKind.Radar,
            PowerPointChartSnapshotKind.Pie => OfficeChartKind.Pie,
            PowerPointChartSnapshotKind.Doughnut => OfficeChartKind.Doughnut,
            PowerPointChartSnapshotKind.Bubble => OfficeChartKind.Bubble,
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported chart snapshot kind.")
        };
    }
}
