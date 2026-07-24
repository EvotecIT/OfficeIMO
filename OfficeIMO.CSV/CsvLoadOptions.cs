#nullable enable

using System.Globalization;
using System.Text;
using System.Threading;

namespace OfficeIMO.CSV;

/// <summary>
/// Options controlling how CSV content is parsed and loaded.
/// </summary>
public sealed class CsvLoadOptions
{
    private const char DefaultDelimiter = ',';

    /// <summary>Default maximum complete stream input size (256 MiB).</summary>
    public const long DefaultMaxInputBytes = 256L * 1024L * 1024L;

    /// <summary>Default maximum physical lines inspected while deciding whether a quoted comment continues.</summary>
    public const int DefaultMaxCommentContinuationLines = 64;

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
    /// Gets or sets the maximum number of physical continuation lines inspected for a quoted comment record.
    /// Once the limit is reached, buffered lines are replayed as ordinary input instead of being retained while
    /// searching for a closing quote. Defaults to <see cref="DefaultMaxCommentContinuationLines"/>.
    /// </summary>
    public int MaxCommentContinuationLines { get; set; } = DefaultMaxCommentContinuationLines;

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
    /// Gets or sets how duplicate header names are handled. Default renames duplicate headers to keep name-based row access unambiguous.
    /// </summary>
    public CsvDuplicateHeaderBehavior DuplicateHeaderBehavior { get; set; } = CsvDuplicateHeaderBehavior.Rename;

    /// <summary>
    /// Gets or sets a token that should be materialized as <c>null</c> when loading rows into a <see cref="CsvDocument"/>.
    /// Raw string streaming callbacks preserve the source text.
    /// </summary>
    public string? NullValue { get; set; }

    /// <summary>
    /// Gets or sets additional date/time formats used by typed row conversion and schema validation.
    /// </summary>
    public string[]? DateTimeFormats { get; set; }

    /// <summary>
    /// Gets or sets how malformed quoted fields are handled. Default is <see cref="CsvQuoteParsingMode.Lenient"/>
    /// to preserve common PowerShell-style import behavior.
    /// </summary>
    public CsvQuoteParsingMode QuoteParsingMode { get; set; } = CsvQuoteParsingMode.Lenient;

    /// <summary>
    /// Gets or sets columns appended to every loaded row, useful for source file names, import timestamps, or batch metadata.
    /// </summary>
    public IReadOnlyDictionary<string, object?>? StaticColumns { get; set; }

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
    /// Gets or sets the field delimiter text. Leave unset to use <see cref="Delimiter"/>.
    /// Single-character values keep the optimized character delimiter path; longer values enable flexible delimiter parsing.
    /// </summary>
    public string? DelimiterText { get; set; }

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
    /// Defaults to 256 MiB. Set to <c>null</c> only for trusted inputs that intentionally exceed the limit.
    /// </summary>
    public long? MaxDecompressedBytes { get; set; } = DefaultMaxInputBytes;

    /// <summary>
    /// Gets or sets the maximum number of bytes accepted by stream-based and uncompressed path-based load APIs.
    /// Defaults to 256 MiB so untrusted inputs cannot be buffered without bound.
    /// </summary>
    public long MaxInputBytes { get; set; } = DefaultMaxInputBytes;

    /// <summary>
    /// Gets or sets an optional cancellation token checked while reading records.
    /// </summary>
    public CancellationToken CancellationToken { get; set; }

    /// <summary>
    /// Gets or sets how often <see cref="ProgressCallback"/> is called, in emitted records.
    /// Default is <c>0</c>, which disables progress callbacks.
    /// </summary>
    public int ProgressReportInterval { get; set; }

    /// <summary>
    /// Gets or sets an optional callback invoked as records are emitted.
    /// </summary>
    public Action<CsvProgress>? ProgressCallback { get; set; }

    /// <summary>
    /// Gets or sets whether parse errors should be collected into <see cref="ParseErrors"/>.
    /// </summary>
    public bool CollectParseErrors { get; set; }

    /// <summary>
    /// Gets or sets how parse errors are handled. Default throws immediately.
    /// </summary>
    public CsvParseErrorAction ParseErrorAction { get; set; } = CsvParseErrorAction.Throw;

    /// <summary>
    /// Gets or sets the maximum number of collected parse errors before parsing fails. Default is <c>100</c>.
    /// </summary>
    public int MaxParseErrors { get; set; } = 100;

    /// <summary>
    /// Gets or sets the collection receiving parse errors when <see cref="CollectParseErrors"/> is enabled.
    /// </summary>
    public IList<CsvParseError>? ParseErrors { get; set; }

    /// <summary>
    /// Gets or sets an optional maximum length for any parsed field.
    /// </summary>
    public int? MaxFieldLength { get; set; }

    /// <summary>
    /// Gets or sets an optional maximum length for fields parsed from quoted records.
    /// </summary>
    public int? MaxQuotedFieldLength { get; set; }

    /// <summary>
    /// Gets or sets whether curly quote characters are normalized to straight quotes while reading.
    /// </summary>
    public bool NormalizeQuotes { get; set; }

    /// <summary>
    /// Gets or sets whether repeated string values are reused through a per-read string cache.
    /// </summary>
    public bool InternStrings { get; set; }

    /// <summary>
    /// Creates an options copy with mutable collections snapshotted for deferred reads.
    /// </summary>
    public CsvLoadOptions Clone()
    {
        var clone = (CsvLoadOptions)MemberwiseClone();
        clone.Header = Header is null ? null : (string[])Header.Clone();
        clone.DateTimeFormats = DateTimeFormats is null ? null : (string[])DateTimeFormats.Clone();
        clone.DelimiterCandidates = DelimiterCandidates is null ? null : (char[])DelimiterCandidates.Clone();
        // ParseErrors is an output sink rather than configuration. Preserve a caller-provided
        // collection so errors produced through an internal options snapshot remain observable,
        // but do not mutate the source options merely to create a default sink.
        clone.ParseErrors = ParseErrors ?? (CollectParseErrors ? new List<CsvParseError>() : null);
        if (StaticColumns is not null)
        {
            var staticColumns = new Dictionary<string, object?>(StaticColumns.Count);
            foreach (var column in StaticColumns)
            {
                staticColumns.Add(column.Key, column.Value);
            }

            clone.StaticColumns = staticColumns;
        }

        return clone;
    }
}
