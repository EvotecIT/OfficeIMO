using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Threading;
using System.Threading.Tasks;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

public sealed partial class ExportOfficeDocumentPdfCommand {
    private const string ParameterSetDirectory = "Directory";
    private const string ParameterSetFiles = "Files";
    private readonly List<string> _batchInputs = new();

    /// <summary>Source directory for mixed-format PDF export. Batch execution requires PowerShell 7.4 or newer.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDirectory)]
    public string InputDirectory { get; set; } = string.Empty;
    /// <summary>Explicit source paths or FileInfo objects from Get-ChildItem.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetFiles)]
    [Alias("SourcePaths")]
    public object[] InputPaths { get; set; } = Array.Empty<object>();
    /// <summary>PDF destination directory. Source names retain their extension plus .pdf.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDirectory)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetFiles)]
    public string OutputDirectory { get; set; } = string.Empty;
    /// <summary>Optional durable state directory. Completed sources and outputs are hash-verified on restart.</summary>
    [Parameter(ParameterSetName = ParameterSetDirectory)] [Parameter(ParameterSetName = ParameterSetFiles)]
    public string? CheckpointDirectory { get; set; }
    /// <summary>Explicit executable conversion route, for example html-pdf for HTML stored in TXT files. By default the catalog selects each file's PDF route.</summary>
    [Parameter(ParameterSetName = ParameterSetDirectory)] [Parameter(ParameterSetName = ParameterSetFiles)]
    public string? ConversionRouteId { get; set; }
    /// <summary>Concurrent conversions. Default two; bounds keep the batch resource budget explicit.</summary>
    [Parameter(ParameterSetName = ParameterSetDirectory)] [Parameter(ParameterSetName = ParameterSetFiles)]
    [ValidateRange(1, 32)] public int MaximumConcurrency { get; set; } = 2;
    /// <summary>Maximum discovered or explicitly selected source files.</summary>
    [Parameter(ParameterSetName = ParameterSetDirectory)] [Parameter(ParameterSetName = ParameterSetFiles)]
    [ValidateRange(1, int.MaxValue)] public int MaximumFiles { get; set; } = 1_000_000;
    /// <summary>Maximum output bytes per document.</summary>
    [Parameter(ParameterSetName = ParameterSetDirectory)] [Parameter(ParameterSetName = ParameterSetFiles)]
    [ValidateRange(1, long.MaxValue)] public long MaximumOutputBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Restrict directory discovery to these source extensions. Other files are reported as skipped.</summary>
    [Parameter(ParameterSetName = ParameterSetDirectory)] public string[]? SourceExtensions { get; set; }
    /// <summary>Exclude descendants when selecting a directory.</summary>
    [Parameter(ParameterSetName = ParameterSetDirectory)] public SwitchParameter NoRecurse { get; set; }
    /// <summary>Destination conflict policy. Durable checkpoint jobs require Fail.</summary>
    [Parameter(ParameterSetName = ParameterSetDirectory)] [Parameter(ParameterSetName = ParameterSetFiles)]
    [ValidateSet("Fail", "Rename", "Replace")] public string ConflictPolicy { get; set; } = "Fail";
    /// <summary>Retry recorded failures; completed items remain protected.</summary>
    [Parameter(ParameterSetName = ParameterSetDirectory)] [Parameter(ParameterSetName = ParameterSetFiles)]
    public SwitchParameter RetryFailed { get; set; }
    /// <summary>Stream every structured per-file outcome instead of the final summary. Summary mode reports at most 100 detailed failure warnings and one suppressed-count warning.</summary>
    [Parameter(ParameterSetName = ParameterSetDirectory)] [Parameter(ParameterSetName = ParameterSetFiles)]
    public SwitchParameter ItemResults { get; set; }
    /// <summary>Variable receiving the bounded batch summary, including when ItemResults is selected.</summary>
    [Parameter(ParameterSetName = ParameterSetDirectory)] [Parameter(ParameterSetName = ParameterSetFiles)]
    public string? SummaryVariable { get; set; }

    private async Task ExportBatchAsync(string[]? paths) {
#if FRAMEWORK
        await Task.CompletedTask;
        throw new PlatformNotSupportedException("Batch PDF export requires PowerShell 7.4 or newer. Single-file export supports Windows PowerShell 5.1.");
#else
        var request = new OfficeConversionBatchRequest {
            InputDirectory = paths == null ? PdfCommandUtilities.ResolvePath(this, InputDirectory) : null,
            InputPaths = paths,
            ConversionRouteId = ConversionRouteId,
            OutputDirectory = PdfCommandUtilities.ResolvePath(this, OutputDirectory),
            CheckpointDirectory = CheckpointDirectory == null ? null : PdfCommandUtilities.ResolvePath(this, CheckpointDirectory),
            MaximumConcurrency = MaximumConcurrency, MaximumFiles = MaximumFiles,
            MaximumInputBytes = MaximumInputBytes, MaximumOutputBytes = MaximumOutputBytes,
            Recursive = !NoRecurse.IsPresent, SourceExtensions = SourceExtensions, RetryFailed = RetryFailed.IsPresent,
            ConflictPolicy = (OfficeWorkflowConflictPolicy)Enum.Parse(typeof(OfficeWorkflowConflictPolicy), ConflictPolicy, true),
            ConversionOptions = new OfficeWorkflowConversionOptions {
                SourcePassword = Password,
                Word = WordOptions, Excel = ExcelOptions, PowerPoint = PowerPointOptions,
                Markdown = MarkdownOptions, Rtf = RtfOptions, Html = HtmlOptions, PlainText = TextOptions,
                LegacyDocLossPolicy = AllowLegacyImportLoss.IsPresent ? OfficeIMO.OfficeConversionLossPolicy.Allow : OfficeIMO.OfficeConversionLossPolicy.Block
            }
        };
        if (!ShouldProcess(request.OutputDirectory, "Export selected documents to PDF")) return;
        var progress = new BatchProgress(this);
        var result = await new OfficeWorkflowRunner().RunBatchAsync(request, progress, CancelToken);
        if (!ItemResults.IsPresent && progress.Failures > BatchProgress.MaximumFailureWarnings) {
            WriteWarning($"{progress.Failures - BatchProgress.MaximumFailureWarnings} further failure messages were suppressed. Use -ItemResults for structured per-file outcomes.");
        }
        // Session variables are accessed on the captured pipeline context after the await.
        PdfCommandUtilities.SetVariable(this, SummaryVariable, result);
        if (!ItemResults.IsPresent) WriteObject(result);
#endif
    }
#if !FRAMEWORK
    private sealed class BatchProgress(ExportOfficeDocumentPdfCommand command) : IProgress<OfficeConversionBatchItemResult> {
        internal const int MaximumFailureWarnings = 100;
        private int _failures;
        public int Failures => Volatile.Read(ref _failures);
        public void Report(OfficeConversionBatchItemResult item) {
            if (command.ItemResults.IsPresent) {
                command.WriteObject(item);
            } else if (item.Status == OfficeWorkflowStatus.Failed && Interlocked.Increment(ref _failures) <= MaximumFailureWarnings) {
                command.WriteWarning(item.InputPath + ": " + item.Summary);
            }
        }
    }
#endif
}
