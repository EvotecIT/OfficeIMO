namespace OfficeIMO.Rtf;

/// <summary>
/// Section line-numbering settings.
/// </summary>
public sealed class RtfLineNumbering {
    /// <summary>Line-number increment modulus. A value of 0 represents no line numbering.</summary>
    public int? CountBy { get; set; }

    /// <summary>Distance from line numbers to the left text margin in twips.</summary>
    public int? DistanceFromTextTwips { get; set; }

    /// <summary>Beginning line number.</summary>
    public int? StartNumber { get; set; }

    /// <summary>Line-number restart behavior.</summary>
    public RtfLineNumberRestart? Restart { get; set; }

    /// <summary>Sets section line-numbering controls.</summary>
    public RtfLineNumbering Set(
        int? countBy = null,
        int? distanceFromTextTwips = null,
        int? startNumber = null,
        RtfLineNumberRestart? restart = null) {
        ValidateNonNegative(countBy, nameof(countBy));
        ValidateNonNegative(distanceFromTextTwips, nameof(distanceFromTextTwips));
        ValidatePositive(startNumber, nameof(startNumber));
        CountBy = countBy;
        DistanceFromTextTwips = distanceFromTextTwips;
        StartNumber = startNumber;
        Restart = restart;
        return this;
    }

    /// <summary>Whether any line-numbering control has been set.</summary>
    public bool HasAnyValue =>
        CountBy.HasValue ||
        DistanceFromTextTwips.HasValue ||
        StartNumber.HasValue ||
        Restart.HasValue;

    internal void Clear() {
        CountBy = null;
        DistanceFromTextTwips = null;
        StartNumber = null;
        Restart = null;
    }

    private static void ValidateNonNegative(int? value, string parameterName) {
        if (value.HasValue && value.Value < 0) {
            throw new ArgumentOutOfRangeException(parameterName, "Line-numbering value cannot be negative.");
        }
    }

    private static void ValidatePositive(int? value, string parameterName) {
        if (value.HasValue && value.Value <= 0) {
            throw new ArgumentOutOfRangeException(parameterName, "Line-number start must be greater than zero.");
        }
    }
}
