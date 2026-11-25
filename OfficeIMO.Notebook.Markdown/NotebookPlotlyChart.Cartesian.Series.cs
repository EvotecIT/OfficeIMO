using System.Globalization;
using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private static bool TryPaintCartesianSeries(OfficeDrawing drawing, CartesianPanel panel, CartesianBox box,
        CartesianLayout appearance, ref int remainingPoints, CancellationToken cancellationToken, List<NotebookChartMark>? marks,
        List<CartesianLegendItem>? legend) {
        if (box.Width < 105 || box.Height < 82 || panel.X.LockAspect || panel.Y.LockAspect) return false;

        bool bars = panel.Kind == "bar";
        bool webGl = panel.Kind == "scattergl";
        JsonElement first = panel.Traces[0].Source;
        bool horizontal = bars && ValueOrDefault(Property(first, "orientation"), "h", allowMissing: false);
        CartesianAxis categoryAxis = horizontal ? panel.Y : panel.X;
        CartesianAxis valueAxis = horizontal ? panel.X : panel.Y;
        if ((bars && (categoryAxis.Linear || categoryAxis.Minimum.HasValue || valueAxis.Categorical || valueAxis.Reversed)) ||
            (!bars && (panel.X.Categorical || panel.Y.Categorical || panel.X.Reversed || panel.Y.Reversed))) return false;
        JsonElement barMode = Property(appearance.Source, "barmode");
        if (bars && !ValueOrDefault(barMode, "group") &&
            !(panel.Traces.Count == 1 && ValueOrDefault(barMode, "relative", allowMissing: false))) return false;

        string? orientation = bars ? (horizontal ? "h" : "v") : null;
        var series = new List<OfficeSeries>(panel.Traces.Count);
        string[]? categories = null;
        HashSet<string>? offsetGroups = bars && panel.Traces.Count > 1 ? new HashSet<string>(StringComparer.Ordinal) : null;
        int panelPoints = Math.Min(webGl ? MaxWebGlPoints : MaxPoints, remainingPoints);
        foreach (CartesianTrace trace in panel.Traces) {
            cancellationToken.ThrowIfCancellationRequested();
            if (offsetGroups != null && HasProperty(trace.Source, "offsetgroup") &&
                (!TryReadString(Property(trace.Source, "offsetgroup"), MaxLabelLength, out string? group) ||
                    (group!.Length != 0 && !offsetGroups.Add(group)))) return false;
            if (!TryReadTrace(trace.Source, panel.Kind, orientation, categoryAxis.Categorical, trace.Index, panelPoints,
                ref categories, out OfficeSeries? parsed, panel.XId, panel.YId, allowNumericCategories: true)) return false;
            series.Add(parsed!);
            panelPoints -= parsed!.Values.Count;
            remainingPoints -= parsed.Values.Count;
        }

        var data = new OfficeChartData(categories ?? new[] { "0" }, series);
        OfficeChartType type = bars ? (horizontal ? OfficeChartType.Bar : OfficeChartType.Column) : OfficeChartType.Scatter;
        (double? xMin, double? xMax) = CartesianSeriesBounds(panel.X, series, bars, horizontal: true);
        (double? yMin, double? yMax) = CartesianSeriesBounds(panel.Y, series, bars, horizontal: false);
        var layout = new OfficeChartLayout(showLegend: false,
            axisLabelFontSize: 8, axisTitleFontSize: 9, categoryAxisTitle: categoryAxis.Title,
            valueAxisTitle: valueAxis.Title, showCategoryAxisLabels: categoryAxis.ShowLabels,
            showValueAxisLabels: valueAxis.ShowLabels, reverseCategoryAxis: categoryAxis.Reversed,
            horizontalAxisMinimum: xMin, horizontalAxisMaximum: xMax,
            verticalAxisMinimum: yMin, verticalAxisMaximum: yMax,
            horizontalAxisCrossingPosition: OfficeChartAxisCrossingPosition.Minimum,
            verticalAxisCrossingPosition: OfficeChartAxisCrossingPosition.Minimum);
        OfficeChartStyle style = appearance.Style.WithPalette(panel.Traces.Select(trace => appearance.Style.GetSeriesColor(trace.Index)));
        if (legend != null) {
            for (int index = 0; index < series.Count; index++) {
                OfficeSeries item = series[index];
                if (!item.ShowInLegend) continue;
                OfficeColor color = item.Color ?? style.GetSeriesColor(index);
                legend.Add(new CartesianLegendItem(panel.Traces[index].Index, item.Name, color,
                    item.SeriesStyle?.OutlineColor ?? color, bars, item.ConnectLine, item.ShowMarkers, item.MarkerShape));
            }
        }
        var chart = new OfficeChartSnapshot("Plotly subplot", panel.Title, type, data, OfficeLength.Points(box.Width),
            OfficeLength.Points(box.Height), style, layout);
        OfficeDrawing plot = NotebookMarkdownProjection.RenderChartSnapshot(chart, marks, useMinimumCanvas: false,
            offsetX: box.Left, offsetY: box.Top);
        drawing.AddDrawing(plot, box.Left, box.Top);
        return true;
    }

    private static (double? Min, double? Max) CartesianSeriesBounds(CartesianAxis axis, List<OfficeSeries> series,
        bool bars, bool horizontal) {
        // Grouped bars need a zero baseline. Scatter autorange follows the data and reserves
        // enough space beyond the outer points for their symbols.
        if (bars || axis.Minimum.HasValue) return (axis.Minimum, axis.Maximum);
        double min = double.PositiveInfinity;
        double max = double.NegativeInfinity;
        foreach (OfficeSeries item in series) {
            IReadOnlyList<double> values = horizontal ? item.XValues! : item.Values;
            foreach (double value in values) {
                min = Math.Min(min, value);
                max = Math.Max(max, value);
            }
        }
        double padding = max > min ? Math.Max(1e-12, (max - min) * 0.07) : Math.Max(1, Math.Abs(min) * 0.05);
        return (min - padding, max + padding);
    }

    private static bool TryReadNumericCategories(JsonElement source, int maxCount, out string[]? categories) {
        categories = null;
        double[]? values = ReadNumbers(source, maxCount);
        if (values == null) return false;
        if (values.Length > 1) {
            double step = values[1] - values[0];
            if (step <= 0) return false;
            for (int index = 2; index < values.Length; index++) {
                double distance = values[index] - values[index - 1];
                if (Math.Abs(distance - step) > step * 1e-9) return false;
            }
        }
        categories = values.Select(value => value.ToString("G17", CultureInfo.InvariantCulture)).ToArray();
        return true;
    }
}
