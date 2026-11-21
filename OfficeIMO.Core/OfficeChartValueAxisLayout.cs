namespace OfficeIMO.Drawing;

/// <summary>Independent numeric scale and label format for a chart value axis.</summary>
public sealed class OfficeChartValueAxisLayout {
    /// <summary>Creates a linear axis scale with optional explicit bounds, tick units, and numeric formatting.</summary>
    public OfficeChartValueAxisLayout(double? minimum = null, double? maximum = null,
        double? majorUnit = null, double? minorUnit = null, string? numberFormat = null,
        OfficeChartAxisTickMark? majorTickMark = null, OfficeChartAxisTickMark? minorTickMark = null) {
        Minimum = ValidateFinite(minimum, nameof(minimum));
        Maximum = ValidateFinite(maximum, nameof(maximum));
        if (minimum.HasValue && maximum.HasValue && minimum.Value >= maximum.Value)
            throw new ArgumentException("The axis maximum must exceed its minimum.", nameof(maximum));
        MajorUnit = ValidatePositive(majorUnit, nameof(majorUnit));
        MinorUnit = ValidatePositive(minorUnit, nameof(minorUnit));
        MajorTickMark = ValidateTickMark(majorTickMark, nameof(majorTickMark));
        MinorTickMark = ValidateTickMark(minorTickMark, nameof(minorTickMark));
        NumberFormat = string.IsNullOrWhiteSpace(numberFormat) ? null : numberFormat!.Trim();
        if (NumberFormat?.Length > OfficeChartLayout.MaxNumberFormatLength)
            throw new ArgumentOutOfRangeException(nameof(numberFormat), "The axis number format exceeds the supported length.");
    }

    /// <summary>Explicit minimum, or automatic scaling when absent.</summary>
    public double? Minimum { get; private set; }
    /// <summary>Explicit maximum, or automatic scaling when absent.</summary>
    public double? Maximum { get; private set; }
    /// <summary>Major tick spacing, or automatic spacing when absent.</summary>
    public double? MajorUnit { get; private set; }
    /// <summary>Minor tick spacing, or automatic spacing when absent.</summary>
    public double? MinorUnit { get; private set; }
    /// <summary>Numeric label format, or the axis default when absent.</summary>
    public string? NumberFormat { get; }
    /// <summary>Major tick appearance, or the shared layout setting when absent.</summary>
    public OfficeChartAxisTickMark? MajorTickMark { get; }
    /// <summary>Minor tick appearance, or the shared layout setting when absent.</summary>
    public OfficeChartAxisTickMark? MinorTickMark { get; }

    /// <summary>Logarithm base of the axis, or null for a linear scale.</summary>
    /// <remarks>Values that are zero or negative cannot be placed on a logarithmic axis and are not drawn.</remarks>
    public double? LogBase { get; private set; }

    /// <summary>Whether the axis runs from its maximum to its minimum (OOXML maxMin orientation).</summary>
    public bool IsReversed { get; private set; }

    /// <summary>Returns a copy with a logarithmic scale of the given base from 2 to 1000; null makes the axis linear.</summary>
    public OfficeChartValueAxisLayout WithLogarithmicScale(double? logBase) {
        if (logBase.HasValue && (double.IsNaN(logBase.Value) || logBase.Value < 2D || logBase.Value > 1000D))
            throw new ArgumentOutOfRangeException(nameof(logBase), "A logarithmic axis base must be from 2 to 1000.");
        var copy = (OfficeChartValueAxisLayout)MemberwiseClone();
        copy.LogBase = logBase;
        return copy;
    }

    /// <summary>Returns a copy whose axis runs in reverse, from its maximum to its minimum, or forward.</summary>
    public OfficeChartValueAxisLayout WithReversedScale(bool reversed) {
        var copy = (OfficeChartValueAxisLayout)MemberwiseClone();
        copy.IsReversed = reversed;
        return copy;
    }

    /// <summary>Copies the axis with its bounds and units restated in the renderer's linear plotting space.
    /// The log base and direction stay set so labels can be restated back.</summary>
    internal OfficeChartValueAxisLayout WithPlottingScale(double? minimum, double? maximum, double? majorUnit, double? minorUnit) {
        var copy = (OfficeChartValueAxisLayout)MemberwiseClone();
        copy.Minimum = minimum;
        copy.Maximum = maximum;
        copy.MajorUnit = majorUnit;
        copy.MinorUnit = minorUnit;
        return copy;
    }

    /// <summary>Independent secondary value-axis title, when set.</summary>
    public string? Title { get; private set; }

    /// <summary>Whether a title replacement or removal was requested explicitly.</summary>
    public bool IsTitleSpecified { get; private set; }

    /// <summary>Returns a copy with the secondary value-axis title; null or whitespace removes an existing title.</summary>
    public OfficeChartValueAxisLayout WithTitle(string? title) {
        var copy = (OfficeChartValueAxisLayout)MemberwiseClone();
        copy.Title = string.IsNullOrWhiteSpace(title) ? null : title;
        copy.IsTitleSpecified = true;
        return copy;
    }

    private static OfficeChartAxisTickMark? ValidateTickMark(OfficeChartAxisTickMark? value, string name) {
        if (value.HasValue && !Enum.IsDefined(typeof(OfficeChartAxisTickMark), value.Value))
            throw new ArgumentOutOfRangeException(name);
        return value;
    }

    private static double? ValidateFinite(double? value, string name) {
        if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value)))
            throw new ArgumentOutOfRangeException(name, "Axis settings must be finite.");
        return value;
    }
    private static double? ValidatePositive(double? value, string name) {
        ValidateFinite(value, name);
        if (value <= 0D) throw new ArgumentOutOfRangeException(name, "Axis tick units must be positive.");
        return value;
    }
}
