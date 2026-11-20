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
