namespace OfficeIMO.Pdf;

/// <summary>Describes how a row column consumes horizontal space.</summary>
public enum PdfColumnWidthUnit {
    /// <summary>Consumes a weighted share of space left after fixed, automatic, and percentage columns.</summary>
    Relative,
    /// <summary>Consumes an exact number of points.</summary>
    Points,
    /// <summary>Consumes the preferred width of its content, constrained by optional bounds.</summary>
    Auto,
    /// <summary>Consumes a percentage of the row's available column area.</summary>
    Percent
}

/// <summary>
/// Immutable row-column sizing value. Use <see cref="Relative"/>, <see cref="Fixed"/>,
/// <see cref="Auto"/>, or <see cref="Percent"/> to make sizing intent explicit.
/// </summary>
public readonly struct PdfColumnWidth {
    private PdfColumnWidth(PdfColumnWidthUnit unit, double value, global::OfficeIMO.OfficeLength minimum, global::OfficeIMO.OfficeLength? maximum) {
        Unit = unit;
        Value = value;
        Minimum = minimum;
        Maximum = maximum;
    }

    /// <summary>Gets the sizing strategy.</summary>
    public PdfColumnWidthUnit Unit { get; }

    /// <summary>Gets the strategy value: weight, points, or percentage. Automatic widths use zero.</summary>
    public double Value { get; }

    /// <summary>Gets the minimum automatic width in points.</summary>
    public global::OfficeIMO.OfficeLength Minimum { get; }

    /// <summary>Gets the optional maximum automatic width in points.</summary>
    public global::OfficeIMO.OfficeLength? Maximum { get; }

    /// <summary>Creates a column that receives a weighted share of remaining width.</summary>
    public static PdfColumnWidth Relative(double weight = 1D) {
        ValidatePositiveFinite(weight, nameof(weight));
        return new PdfColumnWidth(PdfColumnWidthUnit.Relative, weight, global::OfficeIMO.OfficeLength.Points(0D), null);
    }

    /// <summary>Creates a column with an exact width in points.</summary>
    public static PdfColumnWidth Fixed(global::OfficeIMO.OfficeLength points) {
        ValidatePositiveFinite(points.ToPoints(), nameof(points));
        return new PdfColumnWidth(PdfColumnWidthUnit.Points, points.ToPoints(), global::OfficeIMO.OfficeLength.Points(0D), null);
    }

    /// <summary>Creates a content-sized column, optionally constrained in points.</summary>
    public static PdfColumnWidth Auto(global::OfficeIMO.OfficeLength minimum = default, global::OfficeIMO.OfficeLength? maximum = null) {
        ValidateNonNegativeFinite(minimum.ToPoints(), nameof(minimum));
        if (maximum.HasValue) {
            ValidatePositiveFinite(maximum.Value.ToPoints(), nameof(maximum));
            if (maximum.Value.ToPoints() < minimum.ToPoints()) {
                throw new System.ArgumentOutOfRangeException(nameof(maximum), (maximum?.ToPoints()), "Maximum automatic width cannot be smaller than the minimum width.");
            }
        }

        return new PdfColumnWidth(PdfColumnWidthUnit.Auto, 0D, minimum, maximum);
    }

    /// <summary>Creates a column that consumes a percentage of available row width.</summary>
    public static PdfColumnWidth Percent(double percent) {
        ValidatePositiveFinite(percent, nameof(percent));
        if (percent > 100D) {
            throw new System.ArgumentOutOfRangeException(nameof(percent), percent, "Column width cannot exceed 100%.");
        }

        return new PdfColumnWidth(PdfColumnWidthUnit.Percent, percent, global::OfficeIMO.OfficeLength.Points(0D), null);
    }

    internal void Validate(string parameterName) {
        switch (Unit) {
            case PdfColumnWidthUnit.Relative:
            case PdfColumnWidthUnit.Points:
                ValidatePositiveFinite(Value, parameterName);
                break;
            case PdfColumnWidthUnit.Percent:
                ValidatePositiveFinite(Value, parameterName);
                if (Value > 100D) {
                    throw new System.ArgumentOutOfRangeException(parameterName, Value, "Column width cannot exceed 100%.");
                }
                break;
            case PdfColumnWidthUnit.Auto:
                ValidateNonNegativeFinite(Minimum.ToPoints(), parameterName);
                if (Maximum.HasValue && Maximum.Value.ToPoints() < Minimum.ToPoints()) {
                    throw new System.ArgumentOutOfRangeException(parameterName, (Maximum?.ToPoints()), "Maximum automatic width cannot be smaller than the minimum width.");
                }
                break;
            default:
                throw new System.ArgumentOutOfRangeException(parameterName, Unit, "Unsupported PDF row column width unit.");
        }
    }

    private static void ValidatePositiveFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new System.ArgumentOutOfRangeException(parameterName, value, "Column width must be finite and greater than zero.");
        }
    }

    private static void ValidateNonNegativeFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new System.ArgumentOutOfRangeException(parameterName, value, "Column width must be finite and non-negative.");
        }
    }
}
