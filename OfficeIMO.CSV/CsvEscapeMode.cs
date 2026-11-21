namespace OfficeIMO.CSV;

/// <summary>
/// Controls escape sequences inside unquoted fields. Escapes require unquoted fields:
/// <see cref="CsvQuoteParsingMode.None"/> when reading and <see cref="CsvQuoteMode.Never"/> when writing.
/// </summary>
public enum CsvEscapeMode
{
    /// <summary>Fields are written and read literally.</summary>
    None = 0,

    /// <summary>
    /// Backslash escapes as used by PostgreSQL <c>COPY</c> text format and linear TSV:
    /// <c>\t</c> is a tab, <c>\n</c> a line feed, <c>\r</c> a carriage return and <c>\\</c> a backslash.
    /// Writing encodes those four characters. Reading decodes them and keeps any other backslash
    /// sequence, such as PostgreSQL's <c>\N</c> null marker, as literal text.
    /// </summary>
    Backslash = 1
}
