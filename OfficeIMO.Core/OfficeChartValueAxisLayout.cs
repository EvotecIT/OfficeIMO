using System;

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
    public double? Minimum { get; }
    /// <summary>Explicit maximum, or automatic scaling when absent.</summary>
    public double? Maximum { get; }
    /// <summary>Major tick spacing, or automatic spacing when absent.</summary>
    public double? MajorUnit { get; }
    /// <summary>Minor tick spacing, or automatic spacing when absent.</summary>
    public double? MinorUnit { get; }
    /// <summary>Numeric label format, or the axis default when absent.</summary>
    public string? NumberFormat { get; }
    /// <summary>Major tick appearance, or the shared layout setting when absent.</summary>
    public OfficeChartAxisTickMark? MajorTickMark { get; }
    /// <summary>Minor tick appearance, or the shared layout setting when absent.</summary>
    public OfficeChartAxisTickMark? MinorTickMark { get; }

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
