#nullable enable

using System.Globalization;
using System.Text;

namespace OfficeIMO.CSV;

/// <summary>
/// Options controlling how CSV content is serialized.
/// </summary>
public sealed class CsvSaveOptions
{
    private const char DefaultDelimiter = ',';

    /// <summary>
    /// Gets or sets the field delimiter character. Default is <c>,</c>.
    /// </summary>
    public char Delimiter { get; set; } = DefaultDelimiter;

    /// <summary>
    /// Gets or sets the newline sequence written between records. Default is <see cref="Environment.NewLine"/>.
    /// </summary>
    public string NewLine { get; set; } = Environment.NewLine;

    /// <summary>
    /// Gets or sets whether to include the header row when writing. Default is <c>true</c>.
    /// </summary>
    public bool IncludeHeader { get; set; } = true;

    /// <summary>
    /// Gets or sets the culture used for formatting values. Defaults to <see cref="CultureInfo.InvariantCulture"/>.
    /// </summary>
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>
    /// Gets or sets the text encoding used when writing to files. Defaults to UTF-8 without BOM when omitted.
    /// </summary>
    public Encoding? Encoding { get; set; }
}
