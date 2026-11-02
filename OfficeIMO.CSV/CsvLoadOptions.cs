#nullable enable

using System.Globalization;
using System.Text;
using System.Threading;

namespace OfficeIMO.CSV;

/// <summary>Options for loading or parsing an in-memory CSV document.</summary>
public sealed class CsvLoadOptions
{
    internal CsvParsingSettings Parsing { get; set; } = new();

        /// <summary>Default maximum complete stream input size (256 MiB).</summary>
    public const long DefaultMaxInputBytes = CsvParsingSettings.DefaultMaxInputBytes;

    /// <summary>Default maximum physical lines inspected while deciding whether a quoted comment continues.</summary>
    public const int DefaultMaxCommentContinuationLines = CsvParsingSettings.DefaultMaxCommentContinuationLines;

    /// <summary>
    /// Gets or sets whether the first record in the file represents the header row. Default is <c>true</c>.
    /// Ignored when <see cref="Header"/> is provided; in that case the first record is treated as data.
    /// </summary>
    public bool HasHeaderRow { get => Parsing.HasHeaderRow; set => Parsing.HasHeaderRow = value; }

    /// <summary>
    /// Gets or sets explicit header names to use for the CSV data. When provided, no input row is consumed as the header.
    /// </summary>
    public string[]? Header { get => Parsing.Header; set => Parsing.Header = value; }

    /// <summary>
    /// Gets or sets the number of parsed records to skip before header discovery or data emission. Default is <c>0</c>.
    /// </summary>
    public int SkipInitialRecords { get => Parsing.SkipInitialRecords; set => Parsing.SkipInitialRecords = value; }

    /// <summary>
    /// Gets or sets whether comment rows are skipped while discovering the header. Default is <c>true</c>.
    /// Comment rows are identified by <see cref="CommentCharacter"/>.
    /// </summary>
    public bool SkipCommentRowsBeforeHeader { get => Parsing.SkipCommentRowsBeforeHeader; set => Parsing.SkipCommentRowsBeforeHeader = value; }

    /// <summary>
    /// Gets or sets whether comment rows are skipped throughout the file. Default is <c>false</c>.
    /// Comment rows are identified by <see cref="CommentCharacter"/>.
    /// </summary>
    public bool SkipCommentRows { get => Parsing.SkipCommentRows; set => Parsing.SkipCommentRows = value; }

    /// <summary>
    /// Gets or sets the maximum number of physical continuation lines inspected for a quoted comment record.
    /// Once the limit is reached, buffered lines are replayed as ordinary input instead of being retained while
    /// searching for a closing quote. Defaults to <see cref="DefaultMaxCommentContinuationLines"/>.
    /// </summary>
    public int MaxCommentContinuationLines { get => Parsing.MaxCommentContinuationLines; set => Parsing.MaxCommentContinuationLines = value; }

    /// <summary>
    /// Gets or sets the character that identifies a comment row when it appears at the start of a record. Default is <c>#</c>.
    /// </summary>
    public char CommentCharacter { get => Parsing.CommentCharacter; set => Parsing.CommentCharacter = value; }

    /// <summary>
    /// Gets or sets whether W3C Extended Log File Format <c>#Fields:</c> rows are recognized as headers. Default is <c>true</c>.
    /// </summary>
    public bool RecognizeW3CFieldsHeader { get => Parsing.RecognizeW3CFieldsHeader; set => Parsing.RecognizeW3CFieldsHeader = value; }

    /// <summary>
    /// Gets or sets whether blank header names are replaced with generated names such as <c>H1</c>. Default is <c>true</c>.
    /// </summary>
    public bool GenerateMissingHeaderNames { get => Parsing.GenerateMissingHeaderNames; set => Parsing.GenerateMissingHeaderNames = value; }

    /// <summary>
    /// Gets or sets how duplicate header names are handled. Default renames duplicate headers to keep name-based row access unambiguous.
    /// </summary>
    public CsvDuplicateHeaderBehavior DuplicateHeaderBehavior { get => Parsing.DuplicateHeaderBehavior; set => Parsing.DuplicateHeaderBehavior = value; }

    /// <summary>
    /// Gets or sets a token that should be materialized as <c>null</c> when loading rows into a <see cref="CsvDocument"/>.
    /// Raw string streaming callbacks preserve the source text.
    /// </summary>
    public string? NullValue { get => Parsing.NullValue; set => Parsing.NullValue = value; }

    /// <summary>
    /// Gets or sets additional date/time formats used by typed row conversion and schema validation.
    /// </summary>
    public string[]? DateTimeFormats { get => Parsing.DateTimeFormats; set => Parsing.DateTimeFormats = value; }

    /// <summary>
    /// Gets or sets whether schema and row-mapping failures may include source values.
    /// Defaults to <see cref="DataMappingErrorValuePolicy.Include"/> for compatibility.
    /// </summary>
    public DataMappingErrorValuePolicy MappingErrorValuePolicy { get => Parsing.MappingErrorValuePolicy; set => Parsing.MappingErrorValuePolicy = value; }

    /// <summary>
    /// Gets or sets how malformed quoted fields are handled. Default is <see cref="CsvQuoteParsingMode.Lenient"/>
    /// to preserve common PowerShell-style import behavior.
    /// </summary>
    public CsvQuoteParsingMode QuoteParsingMode { get => Parsing.QuoteParsingMode; set => Parsing.QuoteParsingMode = value; }

    /// <summary>
    /// Gets or sets columns appended to every loaded row, useful for source file names, import timestamps, or batch metadata.
    /// </summary>
    public IReadOnlyDictionary<string, object?>? StaticColumns { get => Parsing.StaticColumns; set => Parsing.StaticColumns = value; }

    /// <summary>
    /// Gets or sets how parsed records are handled when their field count differs from the header count.
    /// Default pads missing fields and ignores extras, matching common PowerShell CSV import behavior.
    /// </summary>
    public CsvColumnCountMismatchPolicy ColumnCountMismatchPolicy { get => Parsing.ColumnCountMismatchPolicy; set => Parsing.ColumnCountMismatchPolicy = value; }

    /// <summary>
    /// Gets or sets the field delimiter character. Default is <c>,</c>.
    /// </summary>
    public char Delimiter { get => Parsing.Delimiter; set => Parsing.Delimiter = value; }

    /// <summary>
    /// Gets or sets the field delimiter text. Leave unset to use <see cref="Delimiter"/>.
    /// Single-character values keep the optimized character delimiter path; longer values enable flexible delimiter parsing.
    /// </summary>
    public string? DelimiterText { get => Parsing.DelimiterText; set => Parsing.DelimiterText = value; }

    /// <summary>
    /// Gets or sets whether the delimiter should be detected from the first meaningful records.
    /// Detection is opt-in and leaves <see cref="Delimiter"/> unchanged when no candidate clearly fits.
    /// </summary>
    public bool DetectDelimiter { get => Parsing.DetectDelimiter; set => Parsing.DetectDelimiter = value; }

    /// <summary>
    /// Gets or sets delimiter candidates used when <see cref="DetectDelimiter"/> is enabled.
    /// Defaults to comma, semicolon, pipe, and tab.
    /// </summary>
    public char[]? DelimiterCandidates { get => Parsing.DelimiterCandidates; set => Parsing.DelimiterCandidates = value; }

    /// <summary>
    /// Gets or sets whether leading and trailing whitespace should be trimmed from unquoted fields. Default is <c>false</c>.
    /// </summary>
    public bool TrimWhitespace { get => Parsing.TrimWhitespace; set => Parsing.TrimWhitespace = value; }

    /// <summary>
    /// Gets or sets the culture used for type conversions when reading. Defaults to <see cref="CultureInfo.InvariantCulture"/>.
    /// </summary>
    public CultureInfo Culture { get => Parsing.Culture; set => Parsing.Culture = value; }

    /// <summary>
    /// Gets or sets whether empty lines should be preserved. Default is <c>false</c> (empty lines are ignored).
    /// </summary>
    public bool AllowEmptyLines { get => Parsing.AllowEmptyLines; set => Parsing.AllowEmptyLines = value; }

    internal CsvLoadMode Mode { get => Parsing.Mode; set => Parsing.Mode = value; }

    /// <summary>
    /// Gets or sets the text encoding to use when reading from files. Defaults to UTF-8 if not provided.
    /// </summary>
    public Encoding? Encoding { get => Parsing.Encoding; set => Parsing.Encoding = value; }

    /// <summary>
    /// Gets or sets whether an input BOM may select the decoding encoding. Default is true.
    /// When false, the configured encoding remains authoritative and its own preamble is
    /// accepted. Use an encoding with exception fallback to reject invalid encoded input.
    /// This applies to byte-based file and stream APIs; text inputs are already decoded.
    /// </summary>
    public bool DetectEncodingFromByteOrderMarks { get => Parsing.DetectEncodingFromByteOrderMarks; set => Parsing.DetectEncodingFromByteOrderMarks = value; }

    /// <summary>
    /// Gets or sets compression used when reading from files. Default infers compression from the file extension.
    /// </summary>
    public CsvCompressionType CompressionType { get => Parsing.CompressionType; set => Parsing.CompressionType = value; }

    /// <summary>
    /// Gets or sets an optional limit for decompressed bytes read from compressed CSV files.
    /// Defaults to 256 MiB. Set to <c>null</c> only for trusted inputs that intentionally exceed the limit.
    /// </summary>
    public long? MaxDecompressedBytes { get => Parsing.MaxDecompressedBytes; set => Parsing.MaxDecompressedBytes = value; }

    /// <summary>
    /// Gets or sets the maximum number of bytes accepted by stream-based and uncompressed path-based load APIs.
    /// Defaults to 256 MiB so untrusted inputs cannot be buffered without bound.
    /// </summary>
    public long MaxInputBytes { get => Parsing.MaxInputBytes; set => Parsing.MaxInputBytes = value; }

    /// <summary>
    /// Gets or sets an optional cancellation token checked while reading records.
    /// </summary>
    public CancellationToken CancellationToken { get => Parsing.CancellationToken; set => Parsing.CancellationToken = value; }

    internal CancellationToken OperationCancellationToken { get => Parsing.OperationCancellationToken; set => Parsing.OperationCancellationToken = value; }

    /// <summary>
    /// Gets or sets how often <see cref="ProgressCallback"/> is called, in emitted records.
    /// Default is <c>0</c>, which disables progress callbacks.
    /// </summary>
    public int ProgressReportInterval { get => Parsing.ProgressReportInterval; set => Parsing.ProgressReportInterval = value; }

    /// <summary>
    /// Gets or sets an optional callback invoked as records are emitted.
    /// </summary>
    public Action<CsvProgress>? ProgressCallback { get => Parsing.ProgressCallback; set => Parsing.ProgressCallback = value; }

    /// <summary>
    /// Gets or sets whether parse errors should be collected into <see cref="ParseErrors"/>.
    /// </summary>
    public bool CollectParseErrors { get => Parsing.CollectParseErrors; set => Parsing.CollectParseErrors = value; }

    /// <summary>
    /// Gets or sets how parse errors are handled. Default throws immediately.
    /// </summary>
    public CsvParseErrorAction ParseErrorAction { get => Parsing.ParseErrorAction; set => Parsing.ParseErrorAction = value; }

    /// <summary>
    /// Gets or sets the maximum number of collected parse errors before parsing fails. Default is <c>100</c>.
    /// </summary>
    public int MaxParseErrors { get => Parsing.MaxParseErrors; set => Parsing.MaxParseErrors = value; }

    /// <summary>
    /// Gets or sets the collection receiving parse errors when <see cref="CollectParseErrors"/> is enabled.
    /// </summary>
    public IList<CsvParseError>? ParseErrors { get => Parsing.ParseErrors; set => Parsing.ParseErrors = value; }

    /// <summary>
    /// Gets or sets an optional maximum length for any parsed field.
    /// </summary>
    public int? MaxFieldLength { get => Parsing.MaxFieldLength; set => Parsing.MaxFieldLength = value; }

    /// <summary>
    /// Gets or sets an optional maximum length for fields parsed from quoted records.
    /// </summary>
    public int? MaxQuotedFieldLength { get => Parsing.MaxQuotedFieldLength; set => Parsing.MaxQuotedFieldLength = value; }

    /// <summary>
    /// Gets or sets whether curly quote characters are normalized to straight quotes while reading.
    /// </summary>
    public bool NormalizeQuotes { get => Parsing.NormalizeQuotes; set => Parsing.NormalizeQuotes = value; }

    /// <summary>
    /// Gets or sets whether repeated string values are reused through a per-read string cache.
    /// </summary>
    public bool InternStrings { get => Parsing.InternStrings; set => Parsing.InternStrings = value; }

    /// <summary>Creates an options copy with mutable parsing collections snapshotted.</summary>
    public CsvLoadOptions Clone()
    {
        var clone = (CsvLoadOptions)MemberwiseClone();
        clone.Parsing = Parsing.Clone();
        return clone;
    }
}
