namespace OfficeIMO.Drawing;

public sealed partial class OfficeChartLayout {
    /// <summary>How line, scatter and radar series show missing values: a break in the line, a line joining the
    /// neighbouring points, or a point at zero. Area series always draw missing values at zero, and bars leave
    /// their slot empty unless this is <see cref="OfficeChartBlankDisplay.Zero"/>.</summary>
    /// <remarks>A missing value is a non-finite entry (for example <see cref="double.NaN"/>) in <see cref="OfficeSeries.Values"/>.</remarks>
    public OfficeChartBlankDisplay BlankDisplay { get; private set; } = OfficeChartBlankDisplay.Gap;

    /// <summary>Value on the vertical axis at which the horizontal axis crosses it, or null to use
    /// <see cref="HorizontalAxisCrossingPosition"/>. When the vertical axis is a category axis the value is a
    /// one-based category number.</summary>
    public double? HorizontalAxisCrossesAt { get; private set; }

    /// <summary>Value on the horizontal axis at which the vertical axis crosses it, or null to use
    /// <see cref="VerticalAxisCrossingPosition"/>. When the horizontal axis is a category axis the value is a
    /// one-based category number.</summary>
    public double? VerticalAxisCrossesAt { get; private set; }

    /// <summary>Whether data labels are drawn for values beyond the value-axis maximum. When false, as Office
    /// writes by default, a point clipped at an explicit maximum has no label.</summary>
    public bool ShowDataLabelsOverMaximum { get; private set; } = true;

    /// <summary>Returns a layout copy that shows or hides data labels for values beyond the value-axis maximum.</summary>
    public OfficeChartLayout WithDataLabelsOverMaximum(bool show) {
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.ShowDataLabelsOverMaximum = show;
        return copy;
    }

    /// <summary>Whether a chart with a single bar, column, line, radar or scatter series gives each point its own
    /// palette colour and lists the categories in the legend, as Office's "Vary colors by point" does.</summary>
    /// <remarks>Explicit point colours and styles still win. Charts with more than one series ignore the setting.</remarks>
    public bool VaryColorsByPoint { get; private set; }

    /// <summary>Logarithm base of the physical horizontal value axis, or null for a linear scale.</summary>
    /// <remarks>Applies to the value axis of a horizontal bar chart and to the X axis of a scatter chart.
    /// Values that are zero or negative cannot be placed on a logarithmic axis and are not drawn.</remarks>
    public double? HorizontalAxisLogBase { get; private set; }

    /// <summary>Logarithm base of the physical vertical value axis, or null for a linear scale.</summary>
    /// <remarks>Values that are zero or negative cannot be placed on a logarithmic axis and are not drawn.</remarks>
    public double? VerticalAxisLogBase { get; private set; }

    /// <summary>Whether the physical horizontal value axis runs from its maximum on the left to its minimum on the right.</summary>
    public bool ReverseHorizontalValueAxis { get; private set; }

    /// <summary>Whether the physical vertical value axis runs from its maximum at the bottom to its minimum at the top.</summary>
    public bool ReverseVerticalValueAxis { get; private set; }

    /// <summary>Returns a layout copy with logarithmic value axes. Null keeps an axis linear.</summary>
    /// <param name="horizontalBase">Base from 2 to 1000 for the physical horizontal value axis, or null.</param>
    /// <param name="verticalBase">Base from 2 to 1000 for the physical vertical value axis, or null.</param>
    public OfficeChartLayout WithLogarithmicValueAxes(double? horizontalBase, double? verticalBase) {
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.HorizontalAxisLogBase = ValidateLogBase(horizontalBase, nameof(horizontalBase));
        copy.VerticalAxisLogBase = ValidateLogBase(verticalBase, nameof(verticalBase));
        return copy;
    }

    /// <summary>Returns a layout copy whose value axes run in reverse (OOXML maxMin orientation) or forward.</summary>
    public OfficeChartLayout WithReversedValueAxes(bool horizontal, bool vertical) {
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.ReverseHorizontalValueAxis = horizontal;
        copy.ReverseVerticalValueAxis = vertical;
        return copy;
    }

    /// <summary>Copies the layout with one value axis's bounds, units and perpendicular crossing restated in the
    /// renderer's linear plotting space. The log base and direction flags stay set so labels can be restated back.</summary>
    internal OfficeChartLayout WithPlottingScale(bool horizontal, double? minimum, double? maximum,
        double? majorUnit, double? minorUnit, OfficeChartAxisCrossingPosition crossing, double? crossesAt) {
        var copy = (OfficeChartLayout)MemberwiseClone();
        if (horizontal) {
            copy.HorizontalAxisMinimum = minimum; copy.HorizontalAxisMaximum = maximum;
            copy.HorizontalAxisMajorUnit = majorUnit; copy.HorizontalAxisMinorUnit = minorUnit;
            copy.VerticalAxisCrossingPosition = crossing; copy.VerticalAxisCrossesAt = crossesAt;
        } else {
            copy.VerticalAxisMinimum = minimum; copy.VerticalAxisMaximum = maximum;
            copy.VerticalAxisMajorUnit = majorUnit; copy.VerticalAxisMinorUnit = minorUnit;
            copy.HorizontalAxisCrossingPosition = crossing; copy.HorizontalAxisCrossesAt = crossesAt;
        }
        return copy;
    }

    private static double? ValidateLogBase(double? value, string name) {
        if (value.HasValue && (double.IsNaN(value.Value) || value.Value < 2D || value.Value > 1000D))
            throw new ArgumentOutOfRangeException(name, "A logarithmic axis base must be from 2 to 1000.");
        return value;
    }

    /// <summary>Hand-set plot-area rectangle, or null for the automatic layout.</summary>
    public OfficeChartManualLayout? PlotAreaLayout { get; private set; }

    /// <summary>Hand-set chart-title position, or null for the automatic layout.</summary>
    public OfficeChartManualLayout? TitleLayout { get; private set; }

    /// <summary>Hand-set legend position and size, or null for the automatic layout.</summary>
    public OfficeChartManualLayout? LegendLayout { get; private set; }

    /// <summary>Returns a layout copy with hand-set plot-area, title and legend placement. Null keeps an element's automatic layout.</summary>
    public OfficeChartLayout WithManualLayouts(OfficeChartManualLayout? plotArea, OfficeChartManualLayout? title,
        OfficeChartManualLayout? legend) {
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.PlotAreaLayout = plotArea;
        copy.TitleLayout = title;
        copy.LegendLayout = legend;
        return copy;
    }

    /// <summary>Returns a layout copy that varies a single series' colours by point, or stops doing so.</summary>
    public OfficeChartLayout WithVaryColorsByPoint(bool vary) {
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.VaryColorsByPoint = vary;
        return copy;
    }

    /// <summary>Returns a layout copy that shows missing series values the requested way.</summary>
    /// <param name="blankDisplay">How missing values are drawn.</param>
    public OfficeChartLayout WithBlankDisplay(OfficeChartBlankDisplay blankDisplay) {
        if (!Enum.IsDefined(typeof(OfficeChartBlankDisplay), blankDisplay))
            throw new ArgumentOutOfRangeException(nameof(blankDisplay));
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.BlankDisplay = blankDisplay;
        return copy;
    }

    /// <summary>Returns a layout copy that places each physical axis where it crosses the perpendicular axis.</summary>
    /// <param name="horizontalPosition">Where the horizontal axis crosses the vertical axis.</param>
    /// <param name="horizontalCrossesAt">Explicit crossing value on the vertical axis, which overrides <paramref name="horizontalPosition"/>.</param>
    /// <param name="verticalPosition">Where the vertical axis crosses the horizontal axis.</param>
    /// <param name="verticalCrossesAt">Explicit crossing value on the horizontal axis, which overrides <paramref name="verticalPosition"/>.</param>
    public OfficeChartLayout WithAxisCrossings(OfficeChartAxisCrossingPosition horizontalPosition, double? horizontalCrossesAt,
        OfficeChartAxisCrossingPosition verticalPosition, double? verticalCrossesAt) {
        if (!Enum.IsDefined(typeof(OfficeChartAxisCrossingPosition), horizontalPosition))
            throw new ArgumentOutOfRangeException(nameof(horizontalPosition));
        if (!Enum.IsDefined(typeof(OfficeChartAxisCrossingPosition), verticalPosition))
            throw new ArgumentOutOfRangeException(nameof(verticalPosition));
        if (horizontalCrossesAt.HasValue) ValidateFinite(horizontalCrossesAt.Value, nameof(horizontalCrossesAt));
        if (verticalCrossesAt.HasValue) ValidateFinite(verticalCrossesAt.Value, nameof(verticalCrossesAt));
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.HorizontalAxisCrossingPosition = horizontalPosition;
        copy.HorizontalAxisCrossesAt = horizontalCrossesAt;
        copy.VerticalAxisCrossingPosition = verticalPosition;
        copy.VerticalAxisCrossesAt = verticalCrossesAt;
        return copy;
    }

    /// <summary>Returns a layout copy that divides value-axis labels by a display unit and optionally names the unit.</summary>
    /// <param name="horizontalDivisor">Positive divisor for horizontal value-axis labels, or null for none.</param>
    /// <param name="horizontalLabel">Unit label shown beside the horizontal value axis, or null for none.</param>
    /// <param name="verticalDivisor">Positive divisor for vertical value-axis labels, or null for none.</param>
    /// <param name="verticalLabel">Unit label shown beside the vertical value axis, or null for none.</param>
    public OfficeChartLayout WithAxisDisplayUnits(double? horizontalDivisor, string? horizontalLabel,
        double? verticalDivisor, string? verticalLabel) {
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.HorizontalAxisDisplayUnitDivisor = horizontalDivisor.HasValue ? ValidatePositiveFinite(horizontalDivisor.Value, nameof(horizontalDivisor)) : null;
        copy.HorizontalAxisDisplayUnitLabel = string.IsNullOrWhiteSpace(horizontalLabel) ? null : horizontalLabel;
        copy.VerticalAxisDisplayUnitDivisor = verticalDivisor.HasValue ? ValidatePositiveFinite(verticalDivisor.Value, nameof(verticalDivisor)) : null;
        copy.VerticalAxisDisplayUnitLabel = string.IsNullOrWhiteSpace(verticalLabel) ? null : verticalLabel;
        return copy;
    }
}
