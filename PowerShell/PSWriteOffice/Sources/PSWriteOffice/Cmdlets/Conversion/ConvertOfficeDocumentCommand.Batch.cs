using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Conversion;

public sealed partial class ConvertOfficeDocumentCommand {
    /// <summary>Portable job definition to run or preview; relative paths resolve against this file.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Job")]
    public string? JobPath { get; set; }

    /// <summary>Preview selected inputs and proposed destinations without output or checkpoint writes.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Job")]
    public SwitchParameter Plan { get; set; }

    /// <summary>New durable run report outside the source, output, checkpoint and job trees.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Job")]
    public string? ReportPath { get; set; }

    /// <summary>Traverse child directories when selecting a folder.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter Recurse { get; set; }

    /// <summary>Select these source extensions from directories.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public string[]? Extension { get; set; }

    /// <summary>Maximum discovered files, including skipped inputs; zero selects the largest supported bound.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(0, int.MaxValue - 1)]
    public int MaximumFiles { get; set; } = 1_000_000;

    /// <summary>Maximum simultaneous conversions.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [ValidateRange(1, 32)]
    public int ThrottleLimit { get; set; } = 4;

    /// <summary>Continue a checkpointed batch, verifying completed output hashes before reuse.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public SwitchParameter Resume { get; set; }

    /// <summary>Retry failed checkpoint items while retaining verified completed outputs.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Job")]
    public SwitchParameter RetryFailed { get; set; }

    /// <summary>Checkpoint manifest file. With Resume, defaults to .officeimo-batch.json in the output folder.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public string? CheckpointPath { get; set; }

    /// <summary>Accept an intentionally empty selection without an error.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "Job")]
    public SwitchParameter AllowEmpty { get; set; }

    /// <summary>Save the portable batch definition instead of executing it; returns its FileInfo.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public string? SaveJob { get; set; }

    /// <summary>Display name stored in the saved definition.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public string? JobName { get; set; }

    /// <summary>Admitted environment-variable reference for source passwords in a saved job.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public string? PasswordVariable { get; set; }

    /// <summary>Separate environment-variable reference for PDF source passwords in a mixed-format job.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public string? PdfPasswordVariable { get; set; }

    private OfficeBatchOptions CaptureBatch() {
        if (ParameterSetName == "Job") {
            string path = OfficePath.ResolveInput(SessionState, JobPath!, literal: true).Single();
            var job = OfficeBatchOptions.Load(path);
            job.RetryFailed |= RetryFailed.IsPresent;
            job.AllowEmpty |= AllowEmpty.IsPresent;
            return job;
        }
        if (string.IsNullOrWhiteSpace(OutputDirectory))
            throw new PSArgumentException("Specify OutputDirectory for this batch.");
        string output = OfficePath.ResolveOutput(SessionState, OutputDirectory!, directory: true);
        string target = OfficeFormatSelection.Parse(To)!.Extension;
        string[] directories = _inputs.Where(Directory.Exists).ToArray();
        if (directories.Length > 1 || directories.Length == 1 && _inputs.Count != 1)
            throw new PSArgumentException("Select one directory or a set of files for a portable batch.");
        string? checkpoint = CheckpointPath == null ? Resume
            ? System.IO.Path.Combine(output, ".officeimo-batch.json") : null
            : OfficePath.ResolveOutput(SessionState, CheckpointPath, directory: false);
        if (RetryFailed && checkpoint == null)
            throw new PSArgumentException("RetryFailed requires Resume, CheckpointPath or a saved checkpoint job.");
        if (checkpoint != null && AutoRename)
            throw new PSArgumentException("Checkpointed batches retain their exact output names and cannot use AutoRename.");
        var batch = new OfficeBatchOptions {
            Name = JobName,
            InputDirectory = directories.SingleOrDefault(),
            InputPaths = directories.Length == 0 ? _inputs.ToArray() : null,
            OutputDirectory = output,
            TargetExtension = target,
            ConversionRouteId = Route,
            ConversionOptions = CaptureInputFormat(),
            OutputProfile = OutputProfile,
            ConflictPolicy = checkpoint == null ? ConflictPolicy : OfficePublicationPolicy.Fail,
            SkipExistingOutputs = NoClobber.IsPresent,
            SkipUpToDate = SkipUpToDate.IsPresent,
            Recursive = Recurse.IsPresent,
            Extensions = Extension,
            ResumeFrom = checkpoint,
            RetryFailed = RetryFailed.IsPresent,
            MaxParallelism = ThrottleLimit,
            MaximumFiles = MaximumFiles == 0 ? int.MaxValue - 1 : MaximumFiles,
            MaximumInputBytes = MaximumInputBytes,
            MaximumOutputBytes = MaximumOutputBytes,
            MaximumXmlCharactersInPart = MaximumXmlCharactersInPart,
            SourcePasswordEnvironmentVariable = PasswordVariable,
            PdfPasswordEnvironmentVariable = PdfPasswordVariable,
            AllowEmpty = AllowEmpty.IsPresent
        };
        var routes = Options == null && Route == null && From == null ? null : OfficeBatchExecutor.SelectCapabilities(batch, CancelToken);
        var conversion = CaptureOptions(_inputs, target, routes);
        batch.ConversionOptions = conversion;
        batch.OutputProfile = conversion.OutputProfile;
        batch.MaximumInputBytes = conversion.MaximumInputBytes;
        batch.MaximumOutputBytes = conversion.MaximumOutputBytes;
        batch.MaximumXmlCharactersInPart = conversion.MaximumXmlCharactersInPart;
        return batch;
    }

    private async Task ConvertBatchAsync() {
        OfficeBatchOptions batch = CaptureBatch();
        // Capture provider paths before the first await. Native planning never resolves secret references.
        string? reportPath = ReportPath == null ? null : OfficePath.ResolveOutput(SessionState, ReportPath, directory: false);
        string? savedPath = SaveJob == null ? null : OfficePath.ResolveOutput(SessionState, SaveJob, directory: false);
        if (Plan && (reportPath != null || savedPath != null))
            throw new PSArgumentException("Plan cannot be combined with ReportPath or SaveJob because it performs no writes.");
        if (savedPath != null) {
            if (reportPath != null) throw new PSArgumentException("SaveJob saves a definition; ReportPath records an executed attempt.");
            if (ShouldProcess(savedPath, "Save conversion job")) {
                batch.Save(savedPath, overwrite: File.Exists(savedPath) && !NoClobber.IsPresent);
                WriteObject(new FileInfo(savedPath));
            }
            return;
        }
        if (PasswordVariable != null || PdfPasswordVariable != null || JobName != null)
            throw new PSArgumentException("JobName and password-variable references require SaveJob.");
        if (Plan) {
            var plan = await OfficeConverter.PlanAsync(batch, cancellationToken: CancelToken);
            SetSummary(plan);
            WriteObject(plan);
            if (!plan.CanRun && !(batch.AllowEmpty && plan.Selected == 0 && plan.Failed == 0))
                WriteError(new ErrorRecord(new InvalidOperationException("The conversion plan requires attention. Inspect its diagnostics before execution."),
                    "OfficeDocumentPlanInvalid", ErrorCategory.InvalidResult, plan));
            return;
        }
        if (!ShouldProcess(batch.OutputDirectory, "Convert document batch")) return;
        batch.ReportPath = reportPath;
        batch.ResolveEnvironmentVariable = Environment.GetEnvironmentVariable;
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(batch, new ItemProgress(this), CancelToken);
        WriteProgress(new ProgressRecord(1, "Convert documents", "Finished") { RecordType = ProgressRecordType.Completed });
        OfficeWorkflowCommandUtilities.WriteSavedJobOutcome(this, report.Summary, report, SummaryVariable,
            batch.AllowEmpty, itemResults: true);
    }

    private sealed class ItemProgress(ConvertOfficeDocumentCommand command) : IProgress<OfficeBatchItemReport> {
        private int _processed;
        public void Report(OfficeBatchItemReport item) {
            // Retain failures and skipped observations as well as successful conversion reports.
            command.WriteObject(item.Status == OfficeBatchItemStatus.Completed && item.Report != null ? (object)item.Report : item);
            int processed = System.Threading.Interlocked.Increment(ref _processed);
            if (processed == 1 || processed % 100 == 0)
                command.WriteProgress(new ProgressRecord(1, "Convert documents", $"Processed {processed} files"));
        }
    }
}
