#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Controls when CSV fields are quoted while writing.
/// </summary>
public enum CsvQuoteMode
{
    /// <summary>
    /// Quote only fields that require quoting because they contain a quote, newline, carriage return, or delimiter.
    /// </summary>
    AsNeeded = 0,

    /// <summary>
    /// Quote every field.
    /// </summary>
    Always = 1,

    /// <summary>
    /// Never quote fields. Use only for consumers that require non-standard CSV output.
    /// </summary>
    Never = 2
}
