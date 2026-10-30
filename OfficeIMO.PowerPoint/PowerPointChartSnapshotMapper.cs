using System;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>Owns native snapshot projection for public, image, and PDF chart consumers.</summary>
    internal static class PowerPointChartSnapshotMapper {
        internal static OfficeChartSnapshot ToOfficeSnapshot(PowerPointChartSnapshot snapshot,
            double width, double height, OfficeChartStyle? style = null, OfficeChartLayout? layout = null) {
            var series = snapshot.Data.Series.Select(item => {
                // Keep explicit native colors even if they equal the default palette.
                // Caller styles replace only colors materialized for automatic series.
                bool useNativeColor = style == null || !item.HasAutomaticSeriesColor;
                OfficeColor? color = useNativeColor ? item.Color : null;
                return item.BubbleSizes != null
                ? OfficeSeries.CreateBubble(item.Name, item.XValues!, item.Values,
                    item.BubbleSizes, color, item.PointColors,
                    showInLegend: item.ShowInLegend,
                    markerOutlineColor: useNativeColor ? item.StrokeColor ?? color : null,
                    markerOutlineWidth: item.StrokeWidth,
                    showMarkerOutline: item.ShowStroke)
                : new OfficeSeries(item.Name, item.Values, item.XValues, color,
                    pointColors: item.PointColors, showMarkers: item.SharedAppearance?.ShowMarkers ?? true,
                    showInLegend: item.ShowInLegend, connectLine: item.SharedAppearance?.ConnectLine ?? true,
                    markerSize: item.SharedAppearance?.MarkerSize,
                    markerShape: item.SharedAppearance?.MarkerShape,
                    markerOutlineColor: useNativeColor ? item.SharedAppearance?.MarkerOutlineColor ?? item.StrokeColor ?? color : null,
                    markerOutlineWidth: item.SharedAppearance?.MarkerOutlineWidth,
                    strokeWidth: item.StrokeWidth,
                    strokeDashStyle: item.SharedAppearance?.StrokeDashStyle,
                    renderKind: item.ChartKind.HasValue ? MapKind(item.ChartKind.Value) : null,
                    axisGroup: item.AxisGroup);
            }).Select((series, index) =>
                series.WithSeriesStyle(snapshot.Data.Series[index].SharedAppearance?.SeriesStyle).WithPointStyles(snapshot.Data.Series[index].PointStyles)
                    .WithPointExplosions(snapshot.Data.Series[index].SharedAppearance?.PointExplosions)).ToList();
            return new OfficeChartSnapshot(snapshot.Name, snapshot.Title, MapKind(snapshot.ChartKind),
                new OfficeChartData(snapshot.Data.Categories, series), global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height),
                style ?? snapshot.Style, layout ?? snapshot.Layout,
                bubbleScalePercent: snapshot.BubbleScalePercent, bubbleSizeMode: snapshot.BubbleSizeMode,
                radialLayout: snapshot.RadialLayout);
        }

        internal static OfficeChartType MapKind(PowerPointChartSnapshotKind kind) => kind switch {
            PowerPointChartSnapshotKind.ClusteredColumn => OfficeChartType.ColumnClustered,
            PowerPointChartSnapshotKind.StackedColumn => OfficeChartType.ColumnStacked,
            PowerPointChartSnapshotKind.StackedColumn100 => OfficeChartType.ColumnStacked100,
            PowerPointChartSnapshotKind.ClusteredBar => OfficeChartType.BarClustered,
            PowerPointChartSnapshotKind.StackedBar => OfficeChartType.BarStacked,
            PowerPointChartSnapshotKind.StackedBar100 => OfficeChartType.BarStacked100,
            PowerPointChartSnapshotKind.Line => OfficeChartType.Line,
            PowerPointChartSnapshotKind.StackedLine => OfficeChartType.LineStacked,
            PowerPointChartSnapshotKind.StackedLine100 => OfficeChartType.LineStacked100,
            PowerPointChartSnapshotKind.Area => OfficeChartType.Area,
            PowerPointChartSnapshotKind.StackedArea => OfficeChartType.AreaStacked,
            PowerPointChartSnapshotKind.StackedArea100 => OfficeChartType.AreaStacked100,
            PowerPointChartSnapshotKind.Scatter => OfficeChartType.Scatter,
            PowerPointChartSnapshotKind.Radar => OfficeChartType.Radar,
            PowerPointChartSnapshotKind.Pie => OfficeChartType.Pie,
            PowerPointChartSnapshotKind.Doughnut => OfficeChartType.Doughnut,
            PowerPointChartSnapshotKind.Bubble => OfficeChartType.Bubble,
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "Unsupported chart snapshot kind.")
        };
    }
}
