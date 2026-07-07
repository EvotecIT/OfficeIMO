#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Controls how quoted CSV fields are parsed when input does not strictly follow RFC-style quoting.
/// </summary>
public enum CsvQuoteParsingMode
{
    /// <summary>
    /// Preserve the forgiving parser behavior used by PowerShell-style CSV import workflows.
    /// </summary>
    Lenient = 0,

    /// <summary>
    /// Reject malformed quoted fields instead of falling back to permissive parsing.
    /// </summary>
    Strict = 1
}
