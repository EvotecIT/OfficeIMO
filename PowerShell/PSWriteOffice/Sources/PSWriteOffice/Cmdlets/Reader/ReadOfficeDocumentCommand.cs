using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Reader;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Reads supported documents once and projects text, structure or ingestion views.</summary>
/// <example><summary>Read a folder as text.</summary>
/// <code>Read-OfficeDocument ./Reports -Recurse -As Text</code></example>
/// <example><summary>Reuse a read result without opening its source again.</summary>
/// <code>$document = Read-OfficeDocument ./Report.docx
/// $document | Read-OfficeDocument -As Tables</code></example>
[Cmdlet(VerbsCommunications.Read, "OfficeDocument", DefaultParameterSetName = "Path")]
[OutputType(typeof(OfficeDocumentReadResult), typeof(string), typeof(OfficeDocumentPageText), typeof(OfficeDocumentPageMarkdown),
    typeof(ReaderChunk), typeof(ReaderChunkHierarchyResult), typeof(OfficeDocumentStructuredExtractionResult), typeof(ReaderTable),
    typeof(ReaderTableExportBundle), typeof(ReaderVisual), typeof(ReaderVisualExportBundle), typeof(OfficeDocumentAsset), typeof(ReaderIngestResult), typeof(PdfSearchableOcrReview))]
public sealed partial class ReadOfficeDocumentCommand : OfficeInputCmdlet {
    private readonly List<string> _paths = new();
    private readonly Dictionary<string, long> _ocrInputBudgets = new(StringComparer.Ordinal);
    private OfficeReadConfiguration _configuration = null!;
    private OfficeDocumentReader _reader = null!;
    private OfficeOcrOptions? _ocrOptions;
    private IOcrEngine? _ocrEngine;
    private bool _recurseBound, _extensionBound, _maximumFilesBound;
    private long? TotalByteAllowance => MaximumTotalBytes ?? _configuration.Common.FolderOptions?.MaxTotalBytes;
    private int FileCeiling => _maximumFilesBound ? MaximumFiles : Math.Max(1, _configuration.Common.FolderOptions?.MaxFiles ?? MaximumFiles);

    /// <summary>Already read document; its source is never reopened.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "InputObject")]
    public OfficeDocumentReadResult InputObject { get; set; } = null!;
    /// <summary>Representation to return. Document is the default.</summary>
    [Parameter(Position = 1)] public OfficeReadView As { get; set; }
    /// <summary>Return one text or Markdown object for each observed page, slide or worksheet.</summary>
    [Parameter] public SwitchParameter ByPage { get; set; }
    /// <summary>Omit portable page markers from Markdown.</summary>
    [Parameter] public SwitchParameter NoPageMarkers { get; set; }
    /// <summary>Return in-memory table or visual export bundles.</summary>
    [Parameter] public SwitchParameter Bundle { get; set; }
    /// <summary>Indent JSON output.</summary>
    [Parameter] public SwitchParameter Indented { get; set; }
    /// <summary>Omit chunks from an ingestion summary.</summary>
    [Parameter] public SwitchParameter NoChunks { get; set; }
    /// <summary>Hierarchical chunking policy.</summary>
    [Parameter] [OfficeOptionsTransformation(typeof(ReaderHierarchicalChunkingOptions))]
    public ReaderHierarchicalChunkingOptions? HierarchyOptions { get; set; }
    /// <summary>Structured extraction policy.</summary>
    [Parameter] [OfficeOptionsTransformation(typeof(OfficeDocumentStructuredExtractionOptions))]
    public OfficeDocumentStructuredExtractionOptions? StructureOptions { get; set; }

    /// <summary>Spreadsheet worksheet name, one-based position or sheet object.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>A1 range within each selected worksheet.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public string? Range { get; set; }

    /// <summary>Spreadsheet source rows per chunk.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(1, int.MaxValue)]
    public int? ChunkRows { get; set; }

    /// <summary>Treat the first spreadsheet row as data.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter NoHeaderRow { get; set; }

    /// <summary>Exclude Word footnotes.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter NoFootnotes { get; set; }

    /// <summary>Exclude PowerPoint speaker notes.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter NoNotes { get; set; }

    /// <summary>Do not split Markdown at headings.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter NoHeadingChunks { get; set; }

    /// <summary>Do not compute source and chunk hashes.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter NoHashes { get; set; }

    /// <summary>Compute page locations where the native format does not store them.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter IncludePageLocations { get; set; }

    /// <summary>Maximum characters per source chunk.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(256, int.MaxValue)]
    public int? MaximumChunkCharacters { get; set; }

    /// <summary>Maximum rows per table chunk.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(1, int.MaxValue)]
    public int? MaximumTableRows { get; set; }

    /// <summary>Maximum bytes per source. Native owner limits still apply.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(1, long.MaxValue)]
    public long? MaximumInputBytes { get; set; }

    /// <summary>Maximum XML characters in one Office package part.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(1, long.MaxValue)]
    public long? MaximumXmlCharactersInPart { get; set; }

    /// <summary>Maximum items per mail store; zero means all.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(0, int.MaxValue)]
    public int MaximumStoreItems { get; set; } = 1000;

    /// <summary>Recognize text in source images and scanned pages using an optional OCR engine.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter Ocr { get; set; }
    /// <summary>Optional OCR provider and native recognition policies. Implies Ocr.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [OfficeOptionsTransformation(typeof(OfficeOcrOptions))]
    public OfficeOcrOptions? OcrOptions { get; set; }

    /// <summary>Caller-configured immutable Reader.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public OfficeDocumentReader? Reader { get; set; }

    /// <summary>Open password for encrypted sources.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [OfficePasswordTransformation]
    public System.Security.SecureString? Password { get; set; }

    /// <summary>Typed common and format read options, or common/format-keyed hashtables.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [OfficeReadOptionsTransformation]
    public object[]? Options { get; set; }

    /// <summary>Include subfolders.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter Recurse { get; set; }

    /// <summary>Selected source extensions, with or without a leading dot.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    public string[]? Extension { get; set; }

    /// <summary>Maximum selected files; zero means no limit.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(0, int.MaxValue)]
    public int MaximumFiles { get; set; } = 500;

    /// <summary>Maximum combined source bytes accepted in this read batch.</summary>
    [Parameter(ParameterSetName = "Path")] [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(1, long.MaxValue)]
    public long? MaximumTotalBytes { get; set; }

    /// <inheritdoc />
    protected override Task BeginProcessingAsync() {
        _recurseBound = MyInvocation.BoundParameters.ContainsKey(nameof(Recurse));
        _extensionBound = MyInvocation.BoundParameters.ContainsKey(nameof(Extension));
        _maximumFilesBound = MyInvocation.BoundParameters.ContainsKey(nameof(MaximumFiles));
        if (!Enum.IsDefined(typeof(OfficeReadView), As)) throw new PSArgumentException("Unknown document view.");
        if (ByPage && As is not (OfficeReadView.Text or OfficeReadView.Markdown)) throw new PSArgumentException("ByPage applies to Text or Markdown.");
        if (Bundle && As is not (OfficeReadView.Tables or OfficeReadView.Visuals)) throw new PSArgumentException("Bundle applies to Tables or Visuals.");
        if (NoChunks && As != OfficeReadView.Ingest) throw new PSArgumentException("NoChunks applies to Ingest.");
        if (NoPageMarkers && As != OfficeReadView.Markdown) throw new PSArgumentException("NoPageMarkers applies to Markdown.");
        bool recognize = Ocr || OcrOptions != null || As == OfficeReadView.Ocr;
        if (ParameterSetName == "InputObject" && As == OfficeReadView.Ocr)
            throw new PSArgumentException("An OCR review requires an original image or PDF path. Existing normalized results are never reread.");
        if (recognize && Reader != null) throw new PSArgumentException("Use a Reader with its own OCR processor, or Ocr/OcrOptions without a caller-provided Reader.");
        _ocrOptions = recognize ? OfficeReadOcr.Capture(OcrOptions ?? new OfficeOcrOptions()) : null;
        _configuration = BuildConfiguration();
        if (!recognize) _reader = Reader ?? ReaderCommandUtilities.CreateReader(_configuration.Formats, ThrottleLimit);
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override Task ProcessRecordAsync() {
        if (ParameterSetName == "InputObject") {
            if (As == OfficeReadView.Ocr) throw new PSArgumentException("An OCR review requires an original image or PDF path. Existing normalized results are never reread.");
            WriteProjection(InputObject);
        }
        else _paths.AddRange(ResolveInputs(allowDirectory: true));
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override async Task EndProcessingAsync() {
        if (ParameterSetName == "InputObject" || _paths.Count == 0) return;
        if (_ocrOptions != null) {
            _ocrEngine = await OfficeReadOcr.CreateEngineAsync(_ocrOptions, CancelToken);
            var builder = ReaderCommandUtilities.CreateBuilder(_configuration.Formats).WithMaxConcurrentReads(ThrottleLimit);
            if (As != OfficeReadView.Ocr) builder.AddProcessor(new OfficeDocumentOcrProcessor(_ocrEngine, _ocrOptions.ExecutionOptions));
            _reader = builder.Build();
        }
        var batch = new ReaderBatchOptions { MaxDocuments = FileCeiling == 0 ? int.MaxValue : FileCeiling, MaxDegreeOfParallelism = TotalByteAllowance.HasValue ? 1 : ThrottleLimit, MaxTotalBytes = TotalByteAllowance };
        ReaderCommandUtilities.ValidateBatchConcurrency(_reader, batch);
        var folder = OfficeReadConfiguration.CreateFolder(_configuration.Common, _recurseBound ? (bool)Recurse : (bool?)null, Extension, _extensionBound);
        var selected = ReaderCommandUtilities.CollectDocumentPaths(_reader.EnumerateDocumentPaths(_paths, folder, CancelToken), batch.MaxDocuments, folder.Extensions, out bool limited);
        if (limited) WriteWarning($"Document selection reached MaximumFiles ({FileCeiling}). Use zero for no limit.");
        var paths = As == OfficeReadView.Ocr ? SelectWithinByteBudget(selected) : selected;
        if (As == OfficeReadView.Ocr) {
            try {
                await RunBatchAsync(paths, async (path, token) => {
                    var common = _configuration.Common;
                    if (_ocrInputBudgets.TryGetValue(path, out long allowance)) {
                        common = OfficeReadConfiguration.Copy(common);
                        common.MaxInputBytes = Math.Min(common.MaxInputBytes ?? long.MaxValue, allowance);
                    }
                    return await OfficeReadOcr.ReadReviewAsync(path, _ocrEngine!, _ocrOptions!, common,
                        _configuration.Formats.Pdf?.ReadOptions ?? _configuration.Formats.Pdf?.LoadOptions, token);
                }, "Recognize documents");
            } finally { _configuration.Common.OpenPassword = null; }
            return;
        }
        int completed = 0;
        var ingest = As == OfficeReadView.Ingest ? new OfficeReadIngestSummary(NoChunks) : null;
        try {
            await _reader.ReadDocumentsAsCompletedAsync(paths, outcome => {
                CancelToken.ThrowIfCancellationRequested();
                if (ingest != null) ingest.Add(outcome);
                if (!outcome.Succeeded) WriteError(OfficeError.Create(outcome.Error!, "input-unreadable", outcome.Path));
                else if (ingest == null) WriteProjection(outcome.Document!);
                WriteProgress(new ProgressRecord(0, "Read documents", $"{System.Threading.Interlocked.Increment(ref completed)} completed"));
            }, _configuration.Common, batch, CancelToken);
            if (ingest != null) WriteObject(ingest.Complete());
        } finally {
            _configuration.Common.OpenPassword = null;
            WriteProgress(new ProgressRecord(0, "Read documents", "Finished") { RecordType = ProgressRecordType.Completed });
        }
    }

    private IReadOnlyList<string> SelectWithinByteBudget(IReadOnlyList<string> sources) {
        if (!TotalByteAllowance.HasValue) return sources;
        var selected = new List<string>();
        long accepted = 0;
        foreach (string path in sources) {
            CancelToken.ThrowIfCancellationRequested();
            try {
                long length = new System.IO.FileInfo(path).Length;
                if (length > TotalByteAllowance.Value - accepted) {
                    WriteWarning("Skipped because MaximumTotalBytes would be exceeded: " + path);
                    continue;
                }
                accepted += length;
                _ocrInputBudgets.Add(path, length);
                selected.Add(path);
            } catch (Exception error) when (error is System.IO.IOException or UnauthorizedAccessException) {
                WriteError(OfficeError.Create(error, "input-unreadable", path));
            }
        }
        return selected;
    }
}
