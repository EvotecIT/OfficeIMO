using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Converts CSV text to PSCustomObjects or dictionaries.</summary>
/// <para>Reads CSV text from <c>-Text</c> or the pipeline and maps rows by header.</para>
/// <example>
///   <summary>Convert CSV text into rows.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = ConvertFrom-OfficeCsv -Text "Name,Value`nAlpha,1"</code>
///   <para>Parses CSV text and emits row objects without writing a temporary file.</para>
/// </example>
/// <example>
///   <summary>Convert piped CSV lines.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>"Name,Value", "Alpha,1" | ConvertFrom-OfficeCsv</code>
///   <para>Treats piped lines as one CSV stream.</para>
/// </example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeCsv", DefaultParameterSetName = ParameterSetTextDelimiter)]
public sealed class ConvertFromOfficeCsvCommand : PSCmdlet
{
    private const string ParameterSetTextDelimiter = "TextDelimiter";
    private const string ParameterSetTextCulture = "TextCulture";
    private const string ParameterSetTextDetect = "TextDetect";

    private readonly StringBuilder _textInput = new();
    private readonly CsvPowerShellRowWriter _rowWriter = new();
    private bool _asHashtable;
    private bool _hasTextInput;

    /// <summary>CSV text to parse.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetTextCulture)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetTextDetect)]
    public string? Text { get; set; }

    /// <summary>Treat the first record as data and generate default column names.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Explicit header names to use; when provided, the first CSV record is treated as data.</summary>
    [Parameter]
    public string[]? Header { get; set; }

    /// <summary>Number of parsed CSV records to skip before header discovery or data output.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int SkipRows { get; set; }

    /// <summary>Field delimiter character.</summary>
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    public char Delimiter { get; set; } = ',';

    /// <summary>Field delimiter text for multi-character delimiters such as || or ::.</summary>
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    public string? DelimiterText { get; set; }

    /// <summary>Detect the delimiter from the first meaningful records.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter DetectDelimiter { get; set; }

    /// <summary>Delimiter candidates to consider when detecting the delimiter.</summary>
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public char[]? DelimiterCandidates { get; set; }

    /// <summary>Use the list separator from the selected or current culture as the delimiter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextCulture)]
    public SwitchParameter UseCulture { get; set; }

    /// <summary>Trim whitespace around unquoted fields.</summary>
    [Parameter]
    public bool TrimWhitespace { get; set; }

    /// <summary>Allow empty lines in the input.</summary>
    [Parameter]
    public SwitchParameter AllowEmptyLines { get; set; }

    /// <summary>Skip comment rows starting with # while discovering the header.</summary>
    [Parameter]
    public bool SkipCommentRowsBeforeHeader { get; set; } = true;

    /// <summary>Skip comment rows throughout the input.</summary>
    [Parameter]
    public SwitchParameter SkipCommentRows { get; set; }

    /// <summary>Character that identifies comment rows.</summary>
    [Parameter]
    public char CommentCharacter { get; set; } = '#';

    /// <summary>Recognize W3C Extended Log File Format #Fields: rows as headers.</summary>
    [Parameter]
    public bool RecognizeW3CFieldsHeader { get; set; } = true;

    /// <summary>Controls how rows with fewer or more fields than the header are handled.</summary>
    [Parameter]
    public CsvColumnCountMismatchPolicy ColumnCountMismatchPolicy { get; set; } = CsvColumnCountMismatchPolicy.PadMissingFieldsAndIgnoreExtraFields;

    /// <summary>Controls how duplicate header names are handled.</summary>
    [Parameter]
    public CsvDuplicateHeaderBehavior DuplicateHeaderBehavior { get; set; } = CsvDuplicateHeaderBehavior.Rename;

    /// <summary>Token that is materialized as null when converting rows.</summary>
    [Parameter]
    public string? NullValue { get; set; }

    /// <summary>Additional date/time formats used by typed conversions and validation.</summary>
    [Parameter]
    public string[]? DateTimeFormats { get; set; }

    /// <summary>Controls whether malformed quoted fields are parsed leniently or rejected.</summary>
    [Parameter]
    public CsvQuoteParsingMode QuoteParsingMode { get; set; } = CsvQuoteParsingMode.Lenient;

    /// <summary>Static columns appended to every converted row.</summary>
    [Parameter]
    public IDictionary? StaticColumns { get; set; }

    /// <summary>Load mode controlling materialization.</summary>
    [Parameter]
    public CsvReadMode Mode { get; set; } = CsvReadMode.Stream;

    /// <summary>Culture used for type conversions.</summary>
    [Parameter]
    public CultureInfo? Culture { get; set; }

    /// <summary>Emit dictionaries instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        CsvCommandValidation.EnsureHeaderOptions(NoHeader, Header);
        if (DuplicateHeaderBehavior == CsvDuplicateHeaderBehavior.Preserve)
        {
            throw new PSArgumentException("DuplicateHeaderBehavior Preserve cannot be used with row object or hashtable output. Use Rename or Throw.");
        }

        _asHashtable = AsHashtable.IsPresent;
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (_hasTextInput)
        {
            _textInput.AppendLine();
        }

        _textInput.Append(Text ?? string.Empty);
        _hasTextInput = true;
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        ApplyCultureDelimiter();
        var options = CreateLoadOptions();
        var csvText = _textInput.ToString();

        _rowWriter.Reset();
        if (Mode == CsvReadMode.Stream)
        {
            using var reader = CsvDocument.OpenTextDataReader(csvText, options);
            _rowWriter.WriteDataReaderRows(reader, _asHashtable, this);
            return;
        }

        _rowWriter.WriteDocumentRows(CsvDocument.Parse(csvText, options), _asHashtable, this);
    }

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

        Delimiter = separator[0];
    }

    private CsvLoadOptions CreateLoadOptions()
    {
        var options = new CsvLoadOptions
        {
            HasHeaderRow = Header is null && !NoHeader.IsPresent,
            Header = Header,
            SkipInitialRecords = SkipRows,
            Delimiter = Delimiter,
            DelimiterText = DelimiterText,
            DetectDelimiter = DetectDelimiter.IsPresent,
            DelimiterCandidates = DelimiterCandidates,
            TrimWhitespace = TrimWhitespace,
            AllowEmptyLines = AllowEmptyLines.IsPresent,
            SkipCommentRowsBeforeHeader = SkipCommentRowsBeforeHeader,
            SkipCommentRows = SkipCommentRows.IsPresent,
            CommentCharacter = CommentCharacter,
            RecognizeW3CFieldsHeader = RecognizeW3CFieldsHeader,
            DuplicateHeaderBehavior = CsvDuplicateHeaderBehavior.Throw,
            ColumnCountMismatchPolicy = ColumnCountMismatchPolicy
        };

        CsvPowerShellOptionBuilder.ApplyTextLoadOptions(
            options,
            DuplicateHeaderBehavior,
            NullValue,
            DateTimeFormats,
            QuoteParsingMode,
            StaticColumns);

        if (Culture != null)
        {
            options.Culture = Culture;
        }

        return options;
    }

}
