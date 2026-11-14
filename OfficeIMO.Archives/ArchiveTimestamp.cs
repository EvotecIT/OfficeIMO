using System.Globalization;

namespace OfficeIMO.Archives;

/// <summary>Timestamp helpers for reproducible archives.</summary>
public static class ArchiveTimestamp {
    /// <summary>The environment variable defined by the reproducible-builds specification.</summary>
    public const string SourceDateEpochVariable = "SOURCE_DATE_EPOCH";

    /// <summary>Reads SOURCE_DATE_EPOCH as a UTC time, or returns null when the variable is unset or empty.
    /// Assign the result to <see cref="ArchiveWriterOptions.Timestamp"/>, optionally with
    /// <see cref="ArchiveWriterOptions.ClampToTimestamp"/> to only cap newer times.</summary>
    /// <exception cref="FormatException">The variable is set but is not a non-negative decimal count of seconds.</exception>
    public static DateTimeOffset? FromSourceDateEpoch() => ParseSourceDateEpoch(Environment.GetEnvironmentVariable(SourceDateEpochVariable));

    internal static DateTimeOffset? ParseSourceDateEpoch(string? value) {
        if (string.IsNullOrEmpty(value)) return null;
        if (value!.Any(character => character < '0' || character > '9')
            || !long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out long seconds)
            || seconds > 253402300799L)
            throw new FormatException(SourceDateEpochVariable + " must be a non-negative decimal number of seconds since 1970-01-01T00:00:00Z.");
        return DateTimeOffset.FromUnixTimeSeconds(seconds);
    }
}
