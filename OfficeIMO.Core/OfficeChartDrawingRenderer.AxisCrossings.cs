namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    /// <summary>Where each physical axis meets the perpendicular one, as a fraction of the plot from its bottom or left edge.</summary>
    private readonly struct AxisCrossings {
        internal AxisCrossings(double horizontal, double vertical) {
            Horizontal = horizontal;
            Vertical = vertical;
        }

        internal double Horizontal { get; }
        internal double Vertical { get; }
    }

    private static AxisCrossings GetAxisCrossings(OfficeChartSnapshot snapshot, OfficeChartLayout layout, bool barChart,
        ValueRange valueAxisRange) {
        int categoryCount = snapshot.Data.Categories.Count;
        bool onTickMarks = UsesTickMarkCategoryPoints(snapshot, layout);
        double horizontal, vertical;
        if (barChart) {
            // Horizontal bars: the category axis is vertical, so the value axis is positioned by category number.
            horizontal = GetCategoryCrossingRatio(categoryCount, layout.HorizontalAxisCrossingPosition,
                layout.HorizontalAxisCrossesAt, layout.ReverseCategoryAxis, onTickMarks);
            vertical = GetValueCrossingRatio(valueAxisRange, layout.VerticalAxisCrossingPosition, layout.VerticalAxisCrossesAt);
        } else if (IsScatterChart(snapshot.ChartKind)) {
            ValueRange xRange = ApplyValueAxisScale(GetScatterPointRanges(
                GetRenderableScatterSeries(snapshot).Select(item => item.Series).ToList(),
                GetScatterXValues(snapshot.Data.Categories)).XRange, layout, horizontal: true);
            horizontal = GetValueCrossingRatio(valueAxisRange, layout.HorizontalAxisCrossingPosition, layout.HorizontalAxisCrossesAt);
            vertical = GetValueCrossingRatio(xRange, layout.VerticalAxisCrossingPosition, layout.VerticalAxisCrossesAt);
        } else {
            horizontal = GetValueCrossingRatio(valueAxisRange, layout.HorizontalAxisCrossingPosition, layout.HorizontalAxisCrossesAt);
            vertical = GetCategoryCrossingRatio(categoryCount, layout.VerticalAxisCrossingPosition,
                layout.VerticalAxisCrossesAt, layout.ReverseCategoryAxis, onTickMarks);
        }

        return new AxisCrossings(horizontal, vertical);
    }

    private static double GetClusteredBarBaseline(OfficeChartLayout layout, bool horizontalBars, ValueRange range) =>
        horizontalBars
            ? GetValueCrossing(range, layout.VerticalAxisCrossingPosition, layout.VerticalAxisCrossesAt)
            : GetValueCrossing(range, layout.HorizontalAxisCrossingPosition, layout.HorizontalAxisCrossesAt);

    private static double GetValueCrossingRatio(ValueRange range, OfficeChartAxisCrossingPosition position, double? crossesAt) =>
        GetPlotRatio(GetValueCrossing(range, position, crossesAt), range.Min, range.Max);

    // OOXML autoZero crosses at zero when the axis spans it, and otherwise at the end nearest zero.
    private static double GetValueCrossing(ValueRange range, OfficeChartAxisCrossingPosition position, double? crossesAt) {
        if (crossesAt.HasValue) return ClampValueToRange(crossesAt.Value, range.Min, range.Max);
        if (position == OfficeChartAxisCrossingPosition.Maximum) return range.Max;
        if (position == OfficeChartAxisCrossingPosition.Minimum) return range.Min;
        return range.Min > 0D ? range.Min : range.Max < 0D ? range.Max : 0D;
    }

    // On a category axis the minimum and automatic crossing are the first category, which a reversed axis draws last.
    private static double GetCategoryCrossingRatio(int categoryCount, OfficeChartAxisCrossingPosition position,
        double? crossesAt, bool reverse, bool onTickMarks) {
        double ratio;
        if (crossesAt.HasValue && categoryCount > 0) {
            double category = Math.Max(1D, Math.Min(categoryCount, crossesAt.Value));
            ratio = onTickMarks
                ? categoryCount > 1 ? (category - 1D) / (categoryCount - 1D) : 0D
                : (category - 1D) / categoryCount;
        } else {
            ratio = position == OfficeChartAxisCrossingPosition.Maximum ? 1D : 0D;
        }

        return reverse ? 1D - ratio : ratio;
    }
}
