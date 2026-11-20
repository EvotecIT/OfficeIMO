namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    // Office draws an empty cell in an area series at zero whatever the chart's blank setting,
    // and draws it at zero in line, scatter, radar and bar series only when asked to.
    private static OfficeChartSnapshot ApplyBlankDisplay(OfficeChartSnapshot snapshot) {
        bool zeroRequested = snapshot.Layout.BlankDisplay == OfficeChartBlankDisplay.Zero;
        OfficeSeries[]? replaced = null;
        IReadOnlyList<OfficeSeries> series = snapshot.Data.Series;
        for (int s = 0; s < series.Count; s++) {
            OfficeSeries current = series[s];
            OfficeChartType kind = current.RenderKind ?? snapshot.ChartKind;
            bool zero = IsAreaChart(kind) || zeroRequested && !IsPieChart(kind) && !IsDoughnutChart(kind) &&
                kind != OfficeChartType.Bubble && current.BubbleSizes == null;
            if (!zero || !HasMissingValue(current.Values)) continue;
            var values = new double[current.Values.Count];
            for (int i = 0; i < values.Length; i++) values[i] = IsFiniteChartValue(current.Values[i]) ? current.Values[i] : 0D;
            replaced ??= series.ToArray();
            replaced[s] = current.WithValues(values);
        }

        return replaced == null ? snapshot : new OfficeChartSnapshot(snapshot.Name, snapshot.Title, snapshot.ChartKind,
            new OfficeChartData(snapshot.Data.Categories, replaced), snapshot.Width, snapshot.Height, snapshot.Style,
            snapshot.Layout, snapshot.BubbleScalePercent, snapshot.BubbleSizeMode, snapshot.RadialLayout);
    }

    private static bool HasMissingValue(IReadOnlyList<double> values) {
        for (int i = 0; i < values.Count; i++) if (!IsFiniteChartValue(values[i])) return true;
        return false;
    }

    // Spanning a missing value joins the points on either side of it directly, so only plotted points remain.
    private static void CompactPlottedPoints(IReadOnlyList<OfficePoint> points, IReadOnlyList<bool> plotted,
        out OfficePoint[] compactPoints, out bool[] compactPlotted) {
        var kept = new List<OfficePoint>(points.Count);
        for (int i = 0; i < points.Count; i++) if (plotted[i]) kept.Add(points[i]);
        compactPoints = kept.ToArray();
        compactPlotted = Enumerable.Repeat(true, compactPoints.Length).ToArray();
    }
}
