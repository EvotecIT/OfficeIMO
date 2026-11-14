using System.Collections;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Converts CSV text into PSCustomObjects, dictionaries, a DataTable, a data reader or a CsvDocument.</summary>
/// <para>The reverse of ConvertTo-OfficeCsv and the text form of Import-OfficeCsv: piped strings are read as lines of one
/// CSV text, as ConvertFrom-Csv reads them, and parsed with Import-OfficeCsv's options and output shapes. Use
/// Import-OfficeCsv when the source is a file.</para>
/// <example>
///   <summary>Read CSV text as objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>"Name,Total`nAlpha,10`nBeta,7" | ConvertFrom-OfficeCsv</code>
///   <para>Emits one PSCustomObject per row, with property names from the header.</para>
/// </example>
/// <example>
///   <summary>Round-trip objects through CSV text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = Get-Process | Select-Object -First 5 Name, Id | ConvertTo-OfficeCsv | ConvertFrom-OfficeCsv</code>
///   <para>Reads the text ConvertTo-OfficeCsv writes back into objects.</para>
/// </example>
/// <example>
///   <summary>Read semicolon-separated lines as a typed DataTable.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-Content ./export.txt | ConvertFrom-OfficeCsv -Delimiter ';' -As DataTable -InferSchema</code>
///   <para>Joins the lines Get-Content emits and infers column types from the values.</para>
/// </example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeCsv", DefaultParameterSetName = ParameterSetDelimiter)]
[OutputType(typeof(PSObject), typeof(Hashtable), typeof(System.Data.DataTable), typeof(System.Data.IDataReader), typeof(CsvDocument))]
public sealed class ConvertFromOfficeCsvCommand : PSCmdlet {
    private const string ParameterSetDelimiter = "InputObject";
    private const string ParameterSetCulture = "InputObjectCulture";
    private const string ParameterSetDetect = "InputObjectDetect";
    private readonly List<string> _lines = new();
    private ImportOfficeCsvCommand? _import;

    /// <summary>CSV text, as one string or as lines; piped strings are joined with line breaks.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true), AllowEmptyString, AllowEmptyCollection]
    public string[] InputObject { get; set; } = Array.Empty<string>();

    /// <summary>Treat the first record as data and generate default column names.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Explicit header names to use; when provided, the first CSV record is treated as data.</summary>
    [Parameter]
    public string[]? Header { get; set; }

    /// <summary>Number of parsed CSV records to skip before header discovery or data output.</summary>
    [Parameter, ValidateRange(0, int.MaxValue)]
    public int SkipRows { get; set; }

    /// <summary>Field delimiter, including multi-character delimiters.</summary>
    [Parameter(ParameterSetName = ParameterSetDelimiter)]
    public string Delimiter { get; set; } = ",";

    /// <summary>Detect the delimiter from the first meaningful records.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDetect)]
    public SwitchParameter DetectDelimiter { get; set; }

    /// <summary>Delimiter candidates to consider when detecting the delimiter.</summary>
    [Parameter(ParameterSetName = ParameterSetDetect)]
    public char[]? DelimiterCandidates { get; set; }

    /// <summary>Use the list separator from the selected or current culture as the delimiter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetCulture)]
    public SwitchParameter UseCulture { get; set; }

    /// <summary>Trim whitespace around unquoted fields.</summary>
    [Parameter]
    public SwitchParameter TrimWhitespace { get; set; }

    /// <summary>Allow empty lines in the input.</summary>
    [Parameter]
    public SwitchParameter AllowEmptyLines { get; set; }

    /// <summary>Keep comment rows while discovering the header.</summary>
    [Parameter]
    public SwitchParameter NoSkipCommentRowsBeforeHeader { get; set; }

    /// <summary>Skip comment rows throughout the text.</summary>
    [Parameter]
    public SwitchParameter SkipCommentRows { get; set; }

    /// <summary>Character that identifies comment rows.</summary>
    [Parameter]
    public char CommentCharacter { get; set; } = '#';

    /// <summary>Disable W3C Extended Log File Format #Fields: headers.</summary>
    [Parameter]
    public SwitchParameter NoW3CFieldsHeader { get; set; }

    /// <summary>Controls how rows with fewer or more fields than the header are handled.</summary>
    [Parameter]
    public CsvColumnCountMismatchPolicy ColumnCountMismatchPolicy { get; set; } = CsvColumnCountMismatchPolicy.PadMissingFieldsAndIgnoreExtraFields;

    /// <summary>Controls how duplicate header names are handled.</summary>
    [Parameter]
    public CsvDuplicateHeaderBehavior DuplicateHeaderBehavior { get; set; } = CsvDuplicateHeaderBehavior.Rename;

    /// <summary>Token that is materialized as null.</summary>
    [Parameter]
    public string? NullValue { get; set; }

    /// <summary>Additional date/time formats used by typed conversions and validation.</summary>
    [Parameter]
    public string[]? DateTimeFormats { get; set; }

    /// <summary>Controls whether malformed quoted fields are parsed leniently or rejected.</summary>
    [Parameter]
    public CsvQuoteParsingMode QuoteParsingMode { get; set; } = CsvQuoteParsingMode.Lenient;

    /// <summary>Static columns appended to every row.</summary>
    [Parameter]
    public IDictionary? StaticColumns { get; set; }

    /// <summary>Culture used for type conversions, and for the delimiter with UseCulture.</summary>
    [Parameter]
    public CultureInfo? Culture { get; set; }

    /// <summary>How parse errors are handled.</summary>
    [Parameter]
    public CsvParseErrorAction ParseErrorAction { get; set; } = CsvParseErrorAction.Throw;

    /// <summary>Collect parse errors and write them as non-terminating errors after the rows.</summary>
    [Parameter]
    public SwitchParameter CollectParseErrors { get; set; }

    /// <summary>Maximum number of collected parse errors before parsing fails.</summary>
    [Parameter, ValidateRange(0, int.MaxValue)]
    public int MaxParseErrors { get; set; } = 100;

    /// <summary>Maximum length allowed for any parsed field.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaxFieldLength { get; set; }

    /// <summary>Maximum length allowed for fields parsed from quoted records.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaxQuotedFieldLength { get; set; }

    /// <summary>Normalize curly quote characters to straight quotes.</summary>
    [Parameter]
    public SwitchParameter NormalizeQuotes { get; set; }

    /// <summary>Reuse repeated string values through a per-read cache.</summary>
    [Parameter]
    public SwitchParameter InternStrings { get; set; }

    /// <summary>Infer typed columns when -As DataTable or -As DataReader is used.</summary>
    [Parameter]
    public SwitchParameter InferSchema { get; set; }

    /// <summary>Maximum row count inspected when schema inference is enabled.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int SchemaSampleRows { get; set; } = 1000;

    /// <summary>Explicit column types used when emitting DataTable or IDataReader output.</summary>
    [Parameter]
    public IDictionary? ColumnType { get; set; }

    /// <summary>Output shape: Object, Hashtable, DataTable, DataReader or Document.</summary>
    [Parameter]
    public OfficeCsvOutput As { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (string line in InputObject) _lines.Add(line ?? string.Empty);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        // Import-OfficeCsv owns the parsing options and output shapes; this command only collects the text.
        _import = new ImportOfficeCsvCommand();
        foreach (KeyValuePair<string, object> bound in MyInvocation.BoundParameters) {
            if (bound.Key == nameof(InputObject)) continue;
            typeof(ImportOfficeCsvCommand).GetProperty(bound.Key)!.SetValue(_import, bound.Value);
        }
        _import.ImportText(CommandRuntime, string.Join("\n", _lines));
    }

    /// <inheritdoc />
    protected override void StopProcessing() {
        _import?.Cancel();
        base.StopProcessing();
    }
}
