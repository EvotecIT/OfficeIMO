#nullable enable

using System.Globalization;
using System.Text;

namespace OfficeIMO.CSV;

/// <summary>
/// Options controlling how CSV content is parsed and loaded.
/// </summary>
public sealed class CsvLoadOptions
{
    private const char DefaultDelimiter = ',';

    /// <summary>
    /// Gets or sets whether the first record in the file represents the header row. Default is <c>true</c>.
    /// </summary>
    public bool HasHeaderRow { get; set; } = true;

    /// <summary>
    /// Gets or sets the field delimiter character. Default is <c>,</c>.
    /// </summary>
    public char Delimiter { get; set; } = DefaultDelimiter;

    /// <summary>
    /// Gets or sets whether leading and trailing whitespace should be trimmed from unquoted fields. Default is <c>true</c>.
    /// </summary>
    public bool TrimWhitespace { get; set; } = true;

    /// <summary>
    /// Gets or sets the culture used for type conversions when reading. Defaults to <see cref="CultureInfo.InvariantCulture"/>.
    /// </summary>
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>
    /// Gets or sets whether empty lines should be preserved. Default is <c>false</c> (empty lines are ignored).
    /// </summary>
    public bool AllowEmptyLines { get; set; }

    /// <summary>
    /// Gets or sets the load mode controlling materialization behavior. Default is <see cref="CsvLoadMode.InMemory"/>.
    /// Use <see cref="CsvLoadMode.Stream"/> for very large files when you only need forward-only enumeration; prefer InMemory when you plan to sort/filter/transform.
    /// </summary>
    public CsvLoadMode Mode { get; set; } = CsvLoadMode.InMemory;

    /// <summary>
    /// Gets or sets the text encoding to use when reading from files. Defaults to UTF-8 if not provided.
    /// </summary>
    public Encoding? Encoding { get; set; }

    /// <summary>
    /// Creates a shallow copy of the options instance.
    /// </summary>
    public CsvLoadOptions Clone() => (CsvLoadOptions)MemberwiseClone();
}
