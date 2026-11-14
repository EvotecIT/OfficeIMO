using PSWriteOffice.Services;
using System.Collections;
using System.Data;
using System.Globalization;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Imports CSV rows as PSCustomObjects, dictionaries, or a DataTable.</summary>
/// <para>Uses the CSV header to map fields to property names.</para>
/// <example>
///   <summary>Read rows as PSCustomObjects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Import-OfficeCsv -Path .\data.csv | Format-Table</code>
///   <para>Imports each row as a PSCustomObject.</para>
/// </example>
/// <example>
///   <summary>Read rows as dictionaries.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Import-OfficeCsv -Path .\data.csv -As Hashtable | ForEach-Object { $_['Name'] }</code>
///   <para>Uses hashtables for dynamic schemas or key-based access.</para>
/// </example>
/// <example>
///   <summary>Read rows as a DataTable.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Import-OfficeCsv -Path .\data.csv -As DataTable</code>
///   <para>Emits one DataTable per input file for database and table-oriented workflows.</para>
/// </example>
[Cmdlet(VerbsData.Import, "OfficeCsv", DefaultParameterSetName = ParameterSetPathDelimiter)]
public sealed partial class ImportOfficeCsvCommand : PSCmdlet
{
    private const string ParameterSetPathDelimiter = "Path";
    private const string ParameterSetPathCulture = "PathCulture";
    private const string ParameterSetPathDetect = "PathDetect";
    private const string ParameterSetLiteralPathDelimiter = "LiteralPath";
    private const string ParameterSetLiteralPathCulture = "LiteralPathCulture";
    private const string ParameterSetLiteralPathDetect = "LiteralPathDetect";
    private const string ParameterSetTextDelimiter = "Text";
    private const string ParameterSetTextCulture = "TextCulture";
    private const string ParameterSetTextDetect = "TextDetect";
    private const string ParameterSetDocument = "Document";
    private readonly CsvPowerShellRowWriter _rowWriter = new();
    private readonly List<CsvParseError> _parseErrors = new();
    private readonly CancellationTokenSource _cancellation = new();
    private bool _asDataReader;
    private bool _asDataTable;
    private bool _asHashtable;

    /// <summary>CSV document to read when already loaded.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument), ValidateNotNull]
    public CsvDocument? Document { get; set; }

    /// <summary>Path to one or more CSV files. Wildcards are supported.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPathCulture)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPathDetect)]
    public string[]? Path { get; set; }

    /// <summary>Literal path to one or more CSV files.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetLiteralPathDetect)]
    [Alias("PSPath", "FullName")]
    public string[]? LiteralPath { get; set; }

    /// <summary>CSV text to parse without creating a temporary file.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextDetect), AllowEmptyString]
    public string Text { get; set; } = string.Empty;

    /// <summary>Treat the first record as data and generate default column names.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Explicit header names to use; when provided, the first CSV record is treated as data.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public string[]? Header { get; set; }

    /// <summary>Number of parsed CSV records to skip before header discovery or data output.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    [ValidateRange(0, int.MaxValue)]
    public int SkipRows { get; set; }

    /// <summary>Field delimiter, including multi-character delimiters.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    public string Delimiter { get; set; } = ",";

    /// <summary>Detect the delimiter from the first meaningful records.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPathDetect)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter DetectDelimiter { get; set; }

    /// <summary>Delimiter candidates to consider when detecting the delimiter.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public char[]? DelimiterCandidates { get; set; }

    /// <summary>Use the list separator from the selected or current culture as the delimiter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextCulture)]
    public SwitchParameter UseCulture { get; set; }

    /// <summary>Trim whitespace around unquoted fields.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter TrimWhitespace { get; set; }

    /// <summary>Allow empty lines in the input.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter AllowEmptyLines { get; set; }

    /// <summary>Keep comment rows while discovering the header.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter NoSkipCommentRowsBeforeHeader { get; set; }

    /// <summary>Skip comment rows throughout the file.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter SkipCommentRows { get; set; }

    /// <summary>Character that identifies comment rows.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public char CommentCharacter { get; set; } = '#';

    /// <summary>Disable W3C Extended Log File Format #Fields: headers.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter NoW3CFieldsHeader { get; set; }

    /// <summary>Controls how rows with fewer or more fields than the header are handled.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public CsvColumnCountMismatchPolicy ColumnCountMismatchPolicy { get; set; } = CsvColumnCountMismatchPolicy.PadMissingFieldsAndIgnoreExtraFields;

    /// <summary>Controls how duplicate header names are handled.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public CsvDuplicateHeaderBehavior DuplicateHeaderBehavior { get; set; } = CsvDuplicateHeaderBehavior.Rename;

    /// <summary>Token that is materialized as null when importing rows.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public string? NullValue { get; set; }

    /// <summary>Additional date/time formats used by typed conversions and validation.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public string[]? DateTimeFormats { get; set; }

    /// <summary>Controls whether malformed quoted fields are parsed leniently or rejected.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public CsvQuoteParsingMode QuoteParsingMode { get; set; } = CsvQuoteParsingMode.Lenient;

    /// <summary>Static columns appended to every imported row.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public IDictionary? StaticColumns { get; set; }

    /// <summary>Compression used when reading files. Auto infers from the file extension.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public CsvCompressionType CompressionType { get; set; } = CsvCompressionType.Auto;

    /// <summary>Maximum decompressed bytes to read from compressed CSV files.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [ValidateRange(0, long.MaxValue)]
    public long? MaxDecompressedBytes { get; set; }

    /// <summary>
    /// Maximum input bytes to read. Streamed imports read files of any size unless this is set; imports that load the
    /// whole file into memory default to 256 MiB.
    /// </summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [ValidateRange(1, long.MaxValue)]
    public long? MaxInputBytes { get; set; }

    /// <summary>Load mode controlling materialization.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public CsvReadMode Mode { get; set; } = CsvReadMode.Stream;

    /// <summary>Culture used for type conversions.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public CultureInfo? Culture { get; set; }

    /// <summary>Encoding used when reading the file.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public Encoding? Encoding { get; set; }

    /// <summary>How parse errors are handled.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public CsvParseErrorAction ParseErrorAction { get; set; } = CsvParseErrorAction.Throw;

    /// <summary>Collect parse errors and write them as non-terminating errors after each file.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter CollectParseErrors { get; set; }

    /// <summary>Maximum number of collected parse errors before parsing fails.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    [ValidateRange(0, int.MaxValue)]
    public int MaxParseErrors { get; set; } = 100;

    /// <summary>Maximum length allowed for any parsed field.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxFieldLength { get; set; }

    /// <summary>Maximum length allowed for fields parsed from quoted records.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxQuotedFieldLength { get; set; }

    /// <summary>Normalize curly quote characters to straight quotes.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter NormalizeQuotes { get; set; }

    /// <summary>Reuse repeated string values through a per-read cache.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter InternStrings { get; set; }

    /// <summary>Report progress every N parsed records.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    [ValidateRange(1, int.MaxValue)]
    public int? ProgressInterval { get; set; }

    /// <summary>Infer typed columns when -As DataTable or -As DataReader is used.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter InferSchema { get; set; }

    /// <summary>Maximum row count inspected when schema inference is enabled.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    [ValidateRange(1, int.MaxValue)]
    public int SchemaSampleRows { get; set; } = 1000;

    /// <summary>Explicit column types used when emitting DataTable or IDataReader output.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public IDictionary? ColumnType { get; set; }

    /// <summary>Output shape: Object, Hashtable, DataTable, DataReader or Document.</summary>
    [Parameter]
    public OfficeCsvOutput As { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        if (string.IsNullOrEmpty(Delimiter)) throw new PSArgumentException("Delimiter cannot be empty.", nameof(Delimiter));
        CsvCommandValidation.EnsureHeaderOptions(NoHeader, Header);
        if (!Enum.IsDefined(typeof(OfficeCsvOutput), As)) throw new PSArgumentException("Unsupported CSV output shape.", nameof(As));
        if (As != OfficeCsvOutput.DataReader && As != OfficeCsvOutput.Document && DuplicateHeaderBehavior == CsvDuplicateHeaderBehavior.Preserve)
        {
            throw new PSArgumentException("DuplicateHeaderBehavior Preserve cannot be used with row object, hashtable, or DataTable output. Use -As DataReader, or choose Rename or Throw.");
        }

        if (ColumnType is { Count: > 0 })
        {
            if (!(As == OfficeCsvOutput.DataReader) && !(As == OfficeCsvOutput.DataTable))
            {
                throw new PSArgumentException("-ColumnType can only be used with -As DataTable or -As DataReader.");
            }

            if (InferSchema.IsPresent)
            {
                throw new PSArgumentException("Specify either -ColumnType or -InferSchema, not both.");
            }
        }

        _asDataReader = (As == OfficeCsvOutput.DataReader);
        _asDataTable = (As == OfficeCsvOutput.DataTable);
        _asHashtable = (As == OfficeCsvOutput.Hashtable);
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == ParameterSetDocument) {
            _rowWriter.Reset();
            WriteDocumentRows(Document ?? throw new PSArgumentNullException(nameof(Document)));
            return;
        }
        ApplyCultureDelimiter();
        if (ParameterSetName.StartsWith("Text", StringComparison.Ordinal)) {
            WriteText();
            return;
        }
        var document = Document;

        if (document == null)
        {
            ApplyCultureDelimiter();

            foreach (var resolved in ResolveInputPaths())
            {
                if (!File.Exists(resolved))
                {
                    throw new FileNotFoundException($"File '{resolved}' was not found.", resolved);
                }

                if (Mode == CsvReadMode.Stream && As != OfficeCsvOutput.Document && !RequiresMaterializedRows())
                {
                    _rowWriter.Reset();
                    _parseErrors.Clear();
                    var options = CreateReaderInputOptions();
                    if (_asDataReader)
                    {
                        WriteDataReader(resolved, options);
                    }
                    else if (_asDataTable)
                    {
                        WriteDataTable(resolved, options, System.IO.Path.GetFileNameWithoutExtension(resolved));
                    }
                    else
                    {
                        using var reader = CsvDocument.OpenReader(resolved, options: options);
                        _rowWriter.WriteDataReaderRows(reader, _asHashtable, this);
                    }

                    WriteCollectedParseErrors(resolved);

                    continue;
                }

                _parseErrors.Clear();
                document = CsvDocument.Load(resolved, CreateLoadOptions());
                _rowWriter.Reset();
                WriteDocumentRows(document, System.IO.Path.GetFileNameWithoutExtension(resolved));
                WriteCollectedParseErrors(resolved);
            }

            return;
        }

        _rowWriter.Reset();
        WriteDocumentRows(document);
    }

    private void WriteDocumentRows(CsvDocument document, string? tableName = null)
    {
        if (As == OfficeCsvOutput.Document) { WriteObject(document); return; }
        if (_asDataReader)
        {
            WriteDataReader(document);
            return;
        }

        if (_asDataTable)
        {
            WriteDataTable(document, tableName);
            return;
        }

        _rowWriter.WriteDocumentRows(document, _asHashtable, this);
    }

    private void WriteDataTable(CsvDocument document, string? tableName)
    {
        WriteObject(PSObject.AsPSObject(document.ToDataTable(CreateDataTableOptions(tableName))), enumerateCollection: false);
    }

    private void WriteDataTable(string path, CsvReaderOptions options, string? tableName)
    {
        using var reader = CsvDocument.OpenReader(path, ConfigureReaderOptions(options));
        var table = new DataTable(tableName ?? string.Empty);
        table.Load(reader);
        WriteObject(PSObject.AsPSObject(table), enumerateCollection: false);
    }

    private void WriteDataReader(CsvDocument document)
    {
        WriteObject(PSObject.AsPSObject(document.CreateDataReader(CreateDataReaderOptions())), enumerateCollection: false);
    }

    private void WriteDataReader(string path, CsvReaderOptions options)
    {
        WriteObject(PSObject.AsPSObject(CsvDocument.OpenReader(path, ConfigureReaderOptions(options))), enumerateCollection: false);
    }

    /// <inheritdoc />
    protected override void StopProcessing()
    {
        _cancellation.Cancel();
        base.StopProcessing();
    }

    private bool RequiresMaterializedRows() =>
        _asDataTable ||
        (_asDataReader && CollectParseErrors.IsPresent) ||
        (_asDataReader && ProgressInterval.HasValue);

    private void ApplyCultureDelimiter()
    {
        if (!UseCulture.IsPresent)
        {
            return;
        }

        var separator = (Culture ?? CultureInfo.CurrentCulture).TextInfo.ListSeparator;
        if (string.IsNullOrEmpty(separator) || separator.Length != 1)
        {
            throw new PSArgumentException("The selected culture must provide a single-character list separator.");
        }

        Delimiter = separator;
    }

    private bool IsLiteralPathParameterSet() =>
        string.Equals(ParameterSetName, ParameterSetLiteralPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetLiteralPathCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetLiteralPathDetect, StringComparison.OrdinalIgnoreCase);

    private IEnumerable<string> ResolveInputPaths()
    {
        if (IsLiteralPathParameterSet())
        {
            foreach (var literalPath in LiteralPath ?? Array.Empty<string>())
            {
                yield return OfficeFileCommandUtilities.ResolvePath(this, literalPath);
            }

            yield break;
        }

        foreach (var path in Path ?? Array.Empty<string>())
        {
            if (!WildcardPattern.ContainsWildcardCharacters(path))
            {
                yield return OfficeFileCommandUtilities.ResolvePath(this, path);
                continue;
            }

            var resolvedPaths = OfficeFileCommandUtilities.ResolveWildcardPaths(this, path);
            foreach (var resolvedPath in resolvedPaths)
            {
                yield return resolvedPath;
            }
        }
    }

    private CsvLoadOptions CreateLoadOptions()
    {
        var options = new CsvLoadOptions
        {
            HasHeaderRow = Header is null && !NoHeader.IsPresent,
            Header = Header,
            SkipInitialRecords = SkipRows,
            Delimiter = Delimiter.Length == 1 ? Delimiter[0] : ',',
            DelimiterText = DetectDelimiter ? null : Delimiter,
            DetectDelimiter = DetectDelimiter.IsPresent,
            DelimiterCandidates = DelimiterCandidates,
            TrimWhitespace = TrimWhitespace,
            AllowEmptyLines = AllowEmptyLines.IsPresent,
            SkipCommentRowsBeforeHeader = !NoSkipCommentRowsBeforeHeader,
            SkipCommentRows = SkipCommentRows.IsPresent,
            CommentCharacter = CommentCharacter,
            RecognizeW3CFieldsHeader = !NoW3CFieldsHeader,
            ColumnCountMismatchPolicy = ColumnCountMismatchPolicy,
            CancellationToken = _cancellation.Token,
            ProgressReportInterval = ProgressInterval ?? 0,
            ProgressCallback = ProgressInterval.HasValue ? WriteCsvProgress : null,
            ParseErrorAction = ParseErrorAction,
            CollectParseErrors = CollectParseErrors.IsPresent,
            MaxParseErrors = MaxParseErrors,
            ParseErrors = _parseErrors,
            MaxFieldLength = MaxFieldLength,
            MaxQuotedFieldLength = MaxQuotedFieldLength,
            NormalizeQuotes = NormalizeQuotes.IsPresent,
            InternStrings = InternStrings.IsPresent
        };

        CsvPowerShellOptionBuilder.ApplyLoadOptions(
            options,
            DuplicateHeaderBehavior,
            NullValue,
            DateTimeFormats,
            QuoteParsingMode,
            StaticColumns,
            CompressionType,
            MaxDecompressedBytes);

        if (MaxInputBytes.HasValue)
        {
            options.MaxInputBytes = MaxInputBytes.Value;
        }

        if (Culture != null)
        {
            options.Culture = Culture;
        }

        if (Encoding != null)
        {
            options.Encoding = Encoding;
        }

        return options;
    }

    private CsvReaderOptions CreateReaderInputOptions()
    {
        var options = new CsvReaderOptions
        {
            HasHeaderRow = Header is null && !NoHeader.IsPresent,
            Header = Header,
            SkipInitialRecords = SkipRows,
            Delimiter = Delimiter.Length == 1 ? Delimiter[0] : ',',
            DelimiterText = DetectDelimiter ? null : Delimiter,
            DetectDelimiter = DetectDelimiter.IsPresent,
            DelimiterCandidates = DelimiterCandidates,
            TrimWhitespace = TrimWhitespace,
            AllowEmptyLines = AllowEmptyLines.IsPresent,
            SkipCommentRowsBeforeHeader = !NoSkipCommentRowsBeforeHeader,
            SkipCommentRows = SkipCommentRows.IsPresent,
            CommentCharacter = CommentCharacter,
            RecognizeW3CFieldsHeader = !NoW3CFieldsHeader,
            ColumnCountMismatchPolicy = ColumnCountMismatchPolicy,
            CancellationToken = _cancellation.Token,
            ProgressReportInterval = ProgressInterval ?? 0,
            ProgressCallback = ProgressInterval.HasValue ? WriteCsvProgress : null,
            ParseErrorAction = ParseErrorAction,
            CollectParseErrors = CollectParseErrors.IsPresent,
            MaxParseErrors = MaxParseErrors,
            ParseErrors = _parseErrors,
            MaxFieldLength = MaxFieldLength,
            MaxQuotedFieldLength = MaxQuotedFieldLength,
            NormalizeQuotes = NormalizeQuotes.IsPresent,
            InternStrings = InternStrings.IsPresent
        };

        CsvPowerShellOptionBuilder.ApplyLoadOptions(
            options,
            DuplicateHeaderBehavior,
            NullValue,
            DateTimeFormats,
            QuoteParsingMode,
            StaticColumns,
            CompressionType,
            MaxDecompressedBytes);

        if (MaxInputBytes.HasValue)
        {
            options.MaxInputBytes = MaxInputBytes.Value;
        }

        if (Culture != null)
        {
            options.Culture = Culture;
        }

        if (Encoding != null)
        {
            options.Encoding = Encoding;
        }

        return options;
    }

    private void WriteCsvProgress(CsvProgress progress)
    {
        WriteProgress(new ProgressRecord(
            activityId: 1,
            activity: "Importing CSV",
            statusDescription: $"{progress.RecordsRead:N0} records parsed"));
    }

    private void WriteCollectedParseErrors(string path)
    {
        foreach (var error in _parseErrors)
        {
            WriteError(new ErrorRecord(
                error.Exception,
                "CsvParseError",
                ErrorCategory.ParserError,
                path)
            {
                ErrorDetails = new ErrorDetails(error.Message)
            });
        }
    }
}
