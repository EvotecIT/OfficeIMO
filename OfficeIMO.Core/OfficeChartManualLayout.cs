namespace OfficeIMO.Drawing;

/// <summary>A position, and optionally a size, set by hand for a chart's plot area, title or legend
/// (OOXML <c>c:manualLayout</c>). Every value is a fraction of the chart's width or height.</summary>
public sealed class OfficeChartManualLayout {
    /// <summary>Creates a manual layout.</summary>
    /// <param name="x">Left position as a fraction of the chart width: the left edge, or an offset from the automatic position.</param>
    /// <param name="y">Top position as a fraction of the chart height: the top edge, or an offset from the automatic position.</param>
    /// <param name="width">Width as a fraction of the chart width, or the right edge in <see cref="OfficeChartLayoutMode.Edge"/> mode; null keeps the automatic width.</param>
    /// <param name="height">Height as a fraction of the chart height, or the bottom edge in <see cref="OfficeChartLayoutMode.Edge"/> mode; null keeps the automatic height.</param>
    /// <param name="xMode">Whether <paramref name="x"/> is the left edge or an offset from the automatic position.</param>
    /// <param name="yMode">Whether <paramref name="y"/> is the top edge or an offset from the automatic position.</param>
    /// <param name="widthMode">Whether <paramref name="width"/> is a width or the right edge.</param>
    /// <param name="heightMode">Whether <paramref name="height"/> is a height or the bottom edge.</param>
    /// <param name="target">For a plot area, whether the rectangle excludes (inner) or includes (outer) the axis labels.</param>
    public OfficeChartManualLayout(double x, double y, double? width = null, double? height = null,
        OfficeChartLayoutMode xMode = OfficeChartLayoutMode.Factor, OfficeChartLayoutMode yMode = OfficeChartLayoutMode.Factor,
        OfficeChartLayoutMode widthMode = OfficeChartLayoutMode.Factor, OfficeChartLayoutMode heightMode = OfficeChartLayoutMode.Factor,
        OfficeChartLayoutTarget target = OfficeChartLayoutTarget.Outer) {
        X = ValidateFraction(x, nameof(x));
        Y = ValidateFraction(y, nameof(y));
        Width = width.HasValue ? ValidateFraction(width.Value, nameof(width)) : null;
        Height = height.HasValue ? ValidateFraction(height.Value, nameof(height)) : null;
        XMode = ValidateMode(xMode, nameof(xMode));
        YMode = ValidateMode(yMode, nameof(yMode));
        WidthMode = ValidateMode(widthMode, nameof(widthMode));
        HeightMode = ValidateMode(heightMode, nameof(heightMode));
        if (!Enum.IsDefined(typeof(OfficeChartLayoutTarget), target)) throw new ArgumentOutOfRangeException(nameof(target));
        Target = target;
    }

    /// <summary>Left position as a fraction of the chart width.</summary>
    public double X { get; }

    /// <summary>Top position as a fraction of the chart height.</summary>
    public double Y { get; }

    /// <summary>Width (or right edge) as a fraction of the chart width, or null for the automatic width.</summary>
    public double? Width { get; }

    /// <summary>Height (or bottom edge) as a fraction of the chart height, or null for the automatic height.</summary>
    public double? Height { get; }

    /// <summary>How <see cref="X"/> is interpreted.</summary>
    public OfficeChartLayoutMode XMode { get; }

    /// <summary>How <see cref="Y"/> is interpreted.</summary>
    public OfficeChartLayoutMode YMode { get; }

    /// <summary>How <see cref="Width"/> is interpreted.</summary>
    public OfficeChartLayoutMode WidthMode { get; }

    /// <summary>How <see cref="Height"/> is interpreted.</summary>
    public OfficeChartLayoutMode HeightMode { get; }

    /// <summary>For a plot area, whether the rectangle excludes or includes the axis labels.</summary>
    public OfficeChartLayoutTarget Target { get; }

    // Office keeps manual layouts within a few chart widths; anything larger is malformed.
    private static double ValidateFraction(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < -10D || value > 10D)
            throw new ArgumentOutOfRangeException(name, "Manual layout values are fractions of the chart size.");
        return value;
    }

    private static OfficeChartLayoutMode ValidateMode(OfficeChartLayoutMode mode, string name) {
        if (!Enum.IsDefined(typeof(OfficeChartLayoutMode), mode)) throw new ArgumentOutOfRangeException(name);
        return mode;
    }

    // Resolves one dimension against the chart size and the element's automatic placement.
    internal (double Start, double Size) Resolve(bool horizontal, double chartSize, double automaticStart, double automaticSize) {
        double position = horizontal ? X : Y;
        OfficeChartLayoutMode positionMode = horizontal ? XMode : YMode;
        double? size = horizontal ? Width : Height;
        OfficeChartLayoutMode sizeMode = horizontal ? WidthMode : HeightMode;
        double start = positionMode == OfficeChartLayoutMode.Edge ? position * chartSize : automaticStart + position * chartSize;
        double extent = !size.HasValue ? automaticSize
            : sizeMode == OfficeChartLayoutMode.Edge ? size.Value * chartSize - start
            : size.Value * chartSize;
        return (start, Math.Max(1D, extent));
    }
}
