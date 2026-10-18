using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;
using System.Threading;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Loads a CSV document from disk or parses CSV text.</summary>
/// <para>Returns an <see cref="CsvDocument"/> for inspection or further transformations.</para>
/// <example>
///   <summary>Load a CSV file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$csv = Get-OfficeCsv -Path .\data.csv</code>
///   <para>Loads the CSV file into an OfficeIMO CsvDocument.</para>
/// </example>
/// <example>
///   <summary>Parse CSV text with a custom delimiter.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$csv = Get-OfficeCsv -Text \"Name;Total`nAlpha;10\" -Delimiter ';'</code>
///   <para>Parses a semicolon-delimited CSV string into a document.</para>
/// </example>
/// <example>
///   <summary>Inspect headers as a schema hint.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$csv = Get-OfficeCsv -Path .\data.csv; $csv.Header</code>
///   <para>Returns the header list so you can verify the expected column names.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeCsv", DefaultParameterSetName = ParameterSetPathDelimiter)]
[OutputType(typeof(CsvDocument))]
public sealed class GetOfficeCsvCommand : PSCmdlet
{
    private const string ParameterSetPathDelimiter = "PathDelimiter";
    private const string ParameterSetPathCulture = "PathCulture";
    private const string ParameterSetPathDetect = "PathDetect";
    private const string ParameterSetLiteralPathDelimiter = "LiteralPathDelimiter";
    private const string ParameterSetLiteralPathCulture = "LiteralPathCulture";
    private const string ParameterSetLiteralPathDetect = "LiteralPathDetect";
    private const string ParameterSetTextDelimiter = "TextDelimiter";
    private const string ParameterSetTextCulture = "TextCulture";
    private const string ParameterSetTextDetect = "TextDetect";
    private readonly List<CsvParseError> _parseErrors = new();
    private readonly CancellationTokenSource _cancellation = new();

    /// <summary>Path to one or more CSV files. Wildcards are supported.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPathCulture)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPathDetect)]
    [Alias("FilePath", "InputPath")]
    public string[]? Path { get; set; }

    /// <summary>Literal path to one or more CSV files.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetLiteralPathDetect)]
    [Alias("PSPath", "LP")]
    public string[]? LiteralPath { get; set; }

    /// <summary>CSV text to parse.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextDetect)]
    public string Text { get; set; } = string.Empty;

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
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    public char Delimiter { get; set; } = ',';

    /// <summary>Field delimiter text for multi-character delimiters such as || or ::.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    public string? DelimiterText { get; set; }

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
    [Parameter]
    public bool TrimWhitespace { get; set; }

    /// <summary>Allow empty lines in the input.</summary>
    [Parameter]
    public SwitchParameter AllowEmptyLines { get; set; }

    /// <summary>Skip comment rows starting with # while discovering the header.</summary>
    [Parameter]
    public bool SkipCommentRowsBeforeHeader { get; set; } = true;

    /// <summary>Skip comment rows throughout the file.</summary>
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

    /// <summary>Token that is materialized as null when loading rows.</summary>
    [Parameter]
    public string? NullValue { get; set; }

    /// <summary>Additional date/time formats used by typed conversions and validation.</summary>
    [Parameter]
    public string[]? DateTimeFormats { get; set; }

    /// <summary>Controls whether malformed quoted fields are parsed leniently or rejected.</summary>
    [Parameter]
    public CsvQuoteParsingMode QuoteParsingMode { get; set; } = CsvQuoteParsingMode.Lenient;

    /// <summary>Static columns appended to every loaded row.</summary>
    [Parameter]
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

    /// <summary>Culture used for type conversions.</summary>
    [Parameter]
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
    [Parameter]
    public CsvParseErrorAction ParseErrorAction { get; set; } = CsvParseErrorAction.Throw;

    /// <summary>Collect parse errors and write them as non-terminating errors after each input.</summary>
    [Parameter]
    public SwitchParameter CollectParseErrors { get; set; }

    /// <summary>Maximum number of collected parse errors before parsing fails.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int MaxParseErrors { get; set; } = 100;

    /// <summary>Maximum length allowed for any parsed field.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxFieldLength { get; set; }

    /// <summary>Maximum length allowed for fields parsed from quoted records.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxQuotedFieldLength { get; set; }

    /// <summary>Normalize curly quote characters to straight quotes.</summary>
    [Parameter]
    public SwitchParameter NormalizeQuotes { get; set; }

    /// <summary>Reuse repeated string values through a per-read cache.</summary>
    [Parameter]
    public SwitchParameter InternStrings { get; set; }

    /// <summary>Report progress every N parsed records.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? ProgressInterval { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        CsvCommandValidation.EnsureHeaderOptions(NoHeader, Header);

        if (UseCulture.IsPresent)
        {
            var separator = (Culture ?? CultureInfo.CurrentCulture).TextInfo.ListSeparator;
            if (string.IsNullOrEmpty(separator) || separator.Length != 1)
            {
                throw new PSArgumentException("The selected culture must provide a single-character list separator.");
            }

            Delimiter = separator[0];
        }

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
            DuplicateHeaderBehavior = DuplicateHeaderBehavior,
            ColumnCountMismatchPolicy = ColumnCountMismatchPolicy,
            CompressionType = CompressionType,
            MaxDecompressedBytes = MaxDecompressedBytes,
            CancellationToken = _cancellation.Token,
            ProgressReportInterval = ProgressInterval ?? 0,
            ProgressCallback = ProgressInterval.HasValue ? WriteCsvProgress : null,
            QuoteParsingMode = QuoteParsingMode,
            NullValue = NullValue,
            DateTimeFormats = DateTimeFormats,
            ParseErrorAction = ParseErrorAction,
            CollectParseErrors = CollectParseErrors.IsPresent,
            MaxParseErrors = MaxParseErrors,
            ParseErrors = _parseErrors,
            MaxFieldLength = MaxFieldLength,
            MaxQuotedFieldLength = MaxQuotedFieldLength,
            NormalizeQuotes = NormalizeQuotes.IsPresent,
            InternStrings = InternStrings.IsPresent
        };

        if (IsTextParameterSet())
        {
            CsvPowerShellOptionBuilder.ApplyTextLoadOptions(
                options,
                DuplicateHeaderBehavior,
                NullValue,
                DateTimeFormats,
                QuoteParsingMode,
                StaticColumns);
        }
        else
        {
            CsvPowerShellOptionBuilder.ApplyLoadOptions(
                options,
                DuplicateHeaderBehavior,
                NullValue,
                DateTimeFormats,
                QuoteParsingMode,
                StaticColumns,
                CompressionType,
                MaxDecompressedBytes);
        }

        if (Culture != null)
        {
            options.Culture = Culture;
        }

        if (Encoding != null)
        {
            options.Encoding = Encoding;
        }

        if (IsPathParameterSet())
        {
            foreach (var resolvedPath in ResolveInputPaths())
            {
                if (!File.Exists(resolvedPath))
                {
                    throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
                }

                _parseErrors.Clear();
                WriteObject(CsvDocument.Load(resolvedPath, options));
                WriteCollectedParseErrors(resolvedPath);
            }
        }
        else if (IsLiteralPathParameterSet())
        {
            foreach (var resolvedPath in ResolveInputPaths())
            {
                if (!File.Exists(resolvedPath))
                {
                    throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
                }

                _parseErrors.Clear();
                WriteObject(CsvDocument.Load(resolvedPath, options));
                WriteCollectedParseErrors(resolvedPath);
            }
        }
        else
        {
            _parseErrors.Clear();
            WriteObject(CsvDocument.Parse(Text ?? string.Empty, options));
            WriteCollectedParseErrors("Text");
        }
    }

    /// <inheritdoc />
    protected override void StopProcessing()
    {
        _cancellation.Cancel();
        base.StopProcessing();
    }

    private bool IsPathParameterSet() =>
        string.Equals(ParameterSetName, ParameterSetPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetPathCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetPathDetect, StringComparison.OrdinalIgnoreCase);

    private bool IsLiteralPathParameterSet() =>
        string.Equals(ParameterSetName, ParameterSetLiteralPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetLiteralPathCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetLiteralPathDetect, StringComparison.OrdinalIgnoreCase);

    private bool IsTextParameterSet() =>
        string.Equals(ParameterSetName, ParameterSetTextDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetTextCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetTextDetect, StringComparison.OrdinalIgnoreCase);

    private IEnumerable<string> ResolveInputPaths()
    {
        if (IsLiteralPathParameterSet())
        {
            foreach (var literalPath in LiteralPath ?? Array.Empty<string>())
            {
                yield return SessionState.Path.GetUnresolvedProviderPathFromPSPath(literalPath);
            }

            yield break;
        }

        foreach (var path in Path ?? Array.Empty<string>())
        {
            if (!WildcardPattern.ContainsWildcardCharacters(path))
            {
                yield return SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
                continue;
            }

            var resolvedPaths = SessionState.Path.GetResolvedProviderPathFromPSPath(path, out _);
            foreach (var resolvedPath in resolvedPaths)
            {
                yield return resolvedPath;
            }
        }
    }

    private void WriteCsvProgress(CsvProgress progress)
    {
        WriteProgress(new ProgressRecord(
            activityId: 1,
            activity: "Loading CSV",
            statusDescription: $"{progress.RecordsRead:N0} records parsed"));
    }

    private void WriteCollectedParseErrors(string target)
    {
        foreach (var error in _parseErrors)
        {
            WriteError(new ErrorRecord(
                error.Exception,
                "CsvParseError",
                ErrorCategory.ParserError,
                target)
            {
                ErrorDetails = new ErrorDetails(error.Message)
            });
        }
    }
}
