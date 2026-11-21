namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    // Office leaves out the label of a point plotted beyond an explicit axis maximum when
    // showDLblsOverMax is off: the value axis of the point's axis group, and the X axis of a scatter point.
    // Stacked points are judged by the top of their stack, as they are drawn.
    private static HashSet<(int Series, int Point)>? GetLabelsOverMaximum(OfficeChartSnapshot snapshot) {
        OfficeChartLayout layout = snapshot.Layout;
        if (layout.ShowDataLabelsOverMaximum || !layout.ShowDataLabels) return null;
        bool horizontalBars = IsBarChart(snapshot.ChartKind);
        double? primaryMaximum = horizontalBars ? layout.HorizontalAxisMaximum : layout.VerticalAxisMaximum;
        double? secondaryMaximum = layout.SecondaryValueAxis?.Maximum;
        double? scatterXMaximum = IsScatterChart(snapshot.ChartKind) ? layout.HorizontalAxisMaximum : null;
        if (!primaryMaximum.HasValue && !secondaryMaximum.HasValue && !scatterXMaximum.HasValue) return null;

        var hidden = new HashSet<(int Series, int Point)>();
        IReadOnlyList<OfficeSeries> series = snapshot.Data.Series;
        var stacks = new Dictionary<(OfficeChartType Kind, OfficeChartAxisGroup Group), List<int>>();
        for (int s = 0; s < series.Count; s++) {
            OfficeSeries current = series[s];
            OfficeChartType kind = current.RenderKind ?? snapshot.ChartKind;
            double? maximum = current.AxisGroup == OfficeChartAxisGroup.Secondary ? secondaryMaximum : primaryMaximum;
            if (scatterXMaximum is double xMaximum && current.XValues != null) {
                for (int i = 0; i < current.XValues.Count; i++)
                    if (current.XValues[i] > xMaximum) hidden.Add((s, i));
            }
            if (!maximum.HasValue) continue;
            if (IsStackedForLabelMaximum(kind)) {
                if (!stacks.TryGetValue((kind, current.AxisGroup), out List<int>? members))
                    stacks.Add((kind, current.AxisGroup), members = new List<int>());
                members.Add(s);
                continue;
            }
            for (int i = 0; i < current.Values.Count; i++)
                if (current.Values[i] > maximum.Value) hidden.Add((s, i));
        }

        foreach (KeyValuePair<(OfficeChartType Kind, OfficeChartAxisGroup Group), List<int>> stack in stacks) {
            double maximum = (stack.Key.Group == OfficeChartAxisGroup.Secondary ? secondaryMaximum : primaryMaximum)!.Value;
            bool percent = IsPercentStackedBarOrColumnChart(stack.Key.Kind) || IsPercentStackedLineChart(stack.Key.Kind) ||
                IsPercentStackedAreaChart(stack.Key.Kind);
            int count = stack.Value.Max(index => series[index].Values.Count);
            for (int point = 0; point < count; point++) {
                double positiveTotal = 0D, negativeTotal = 0D;
                foreach (int member in stack.Value) {
                    if (!TryGetSeriesValue(series[member], point, out double value)) continue;
                    if (value > 0D) positiveTotal += value;
                    else negativeTotal -= value;
                }
                double positive = 0D, negative = 0D;
                foreach (int member in stack.Value) {
                    if (!TryGetSeriesValue(series[member], point, out double value)) continue;
                    double top;
                    // A zero sits at the top of the positive stack, where its label is drawn.
                    if (value >= 0D) {
                        positive += value;
                        top = percent ? positive / positiveTotal : positive;
                    } else {
                        negative += value;
                        top = percent ? negative / negativeTotal : negative;
                    }
                    if (top > maximum) hidden.Add((member, point));
                }
            }
        }

        return hidden.Count == 0 ? null : hidden;
    }

    private static bool IsStackedForLabelMaximum(OfficeChartType kind) =>
        IsStackedBarOrColumnChart(kind) || IsPercentStackedBarOrColumnChart(kind) ||
        IsStackedLineChart(kind) || IsPercentStackedLineChart(kind) ||
        IsStackedAreaChart(kind) || IsPercentStackedAreaChart(kind);
}
