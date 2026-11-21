namespace OfficeIMO.CSV;

/// <summary>
/// Controls whether quote characters delimit fields and how malformed quoted fields are handled.
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
    Strict = 1,

    /// <summary>
    /// Treat quote characters as ordinary field data. Every line break ends a record, so fields cannot
    /// contain the delimiter or a line break unless <see cref="CsvEscapeMode.Backslash"/> encodes them.
    /// This is the IANA <c>text/tab-separated-values</c> rule used by <see cref="CsvProfile.Tsv"/>.
    /// </summary>
    None = 2
}
