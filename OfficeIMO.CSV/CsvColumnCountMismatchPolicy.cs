#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Controls how parsed CSV records are aligned when the record field count differs from the header count.
/// </summary>
public enum CsvColumnCountMismatchPolicy
{
    /// <summary>
    /// Throw when a parsed record has fewer or more fields than the header.
    /// </summary>
    Strict,

    /// <summary>
    /// Pad missing fields with empty values and ignore fields beyond the header.
    /// </summary>
    PadMissingFieldsAndIgnoreExtraFields
}
