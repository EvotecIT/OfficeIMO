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
    /// Ignored when <see cref="Header"/> is provided; in that case the first record is treated as data.
    /// </summary>
    public bool HasHeaderRow { get; set; } = true;

    /// <summary>
    /// Gets or sets explicit header names to use for the CSV data. When provided, no input row is consumed as the header.
    /// </summary>
    public string[]? Header { get; set; }

    /// <summary>
    /// Gets or sets the number of parsed records to skip before header discovery or data emission. Default is <c>0</c>.
    /// </summary>
    public int SkipInitialRecords { get; set; }

    /// <summary>
    /// Gets or sets whether comment rows are skipped while discovering the header. Default is <c>true</c>.
    /// Comment rows are identified by <see cref="CommentCharacter"/>.
    /// </summary>
    public bool SkipCommentRowsBeforeHeader { get; set; } = true;

    /// <summary>
    /// Gets or sets whether comment rows are skipped throughout the file. Default is <c>false</c>.
    /// Comment rows are identified by <see cref="CommentCharacter"/>.
    /// </summary>
    public bool SkipCommentRows { get; set; }

    /// <summary>
    /// Gets or sets the character that identifies a comment row when it appears at the start of a record. Default is <c>#</c>.
    /// </summary>
    public char CommentCharacter { get; set; } = '#';

    /// <summary>
    /// Gets or sets whether W3C Extended Log File Format <c>#Fields:</c> rows are recognized as headers. Default is <c>true</c>.
    /// </summary>
    public bool RecognizeW3CFieldsHeader { get; set; } = true;

    /// <summary>
    /// Gets or sets whether blank header names are replaced with generated names such as <c>H1</c>. Default is <c>true</c>.
    /// </summary>
    public bool GenerateMissingHeaderNames { get; set; } = true;

    /// <summary>
    /// Gets or sets how parsed records are handled when their field count differs from the header count.
    /// Default pads missing fields and ignores extras, matching common PowerShell CSV import behavior.
    /// </summary>
    public CsvColumnCountMismatchPolicy ColumnCountMismatchPolicy { get; set; } = CsvColumnCountMismatchPolicy.PadMissingFieldsAndIgnoreExtraFields;

    /// <summary>
    /// Gets or sets the field delimiter character. Default is <c>,</c>.
    /// </summary>
    public char Delimiter { get; set; } = DefaultDelimiter;

    /// <summary>
    /// Gets or sets whether the delimiter should be detected from the first meaningful records.
    /// Detection is opt-in and leaves <see cref="Delimiter"/> unchanged when no candidate clearly fits.
    /// </summary>
    public bool DetectDelimiter { get; set; }

    /// <summary>
    /// Gets or sets delimiter candidates used when <see cref="DetectDelimiter"/> is enabled.
    /// Defaults to comma, semicolon, pipe, and tab.
    /// </summary>
    public char[]? DelimiterCandidates { get; set; }

    /// <summary>
    /// Gets or sets whether leading and trailing whitespace should be trimmed from unquoted fields. Default is <c>false</c>.
    /// </summary>
    public bool TrimWhitespace { get; set; }

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
    /// Gets or sets compression used when reading from files. Default infers compression from the file extension.
    /// </summary>
    public CsvCompressionType CompressionType { get; set; } = CsvCompressionType.Auto;

    /// <summary>
    /// Gets or sets an optional limit for decompressed bytes read from compressed CSV files.
    /// </summary>
    public long? MaxDecompressedBytes { get; set; }

    /// <summary>
    /// Creates a shallow copy of the options instance.
    /// </summary>
    public CsvLoadOptions Clone() => (CsvLoadOptions)MemberwiseClone();
}
