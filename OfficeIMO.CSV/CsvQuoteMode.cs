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
    /// Never quote fields. A field containing the delimiter, a carriage return or a line feed cannot be
    /// represented and fails the write, unless <see cref="CsvSaveOptions.EscapeMode"/> is
    /// <see cref="CsvEscapeMode.Backslash"/>. Quote characters are written as ordinary data.
    /// </summary>
    Never = 2
}
