using OfficeIMO;
using System;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using PSWriteOffice.Services;
using OfficeIMO.Workflows;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Plans or runs a saved conversion job through the same runner as Studio and the CLI.</summary>
/// <para>Plan performs no output or checkpoint writes. Execution revalidates current inputs and retains completed artifacts after a partial failure or cancellation. Password environment references are resolved only for execution; values are not written to the job or reports.</para>
/// <example>
///   <summary>Inspect selected files, destinations and prerequisites before executing a saved job.</summary>
///   <code>Invoke-OfficeConversionJob -Path ./monthly.officejob.json -Plan
/// Invoke-OfficeConversionJob -Path ./monthly.officejob.json -ItemResults -SummaryVariable summary -ErrorAction Stop</code>
/// </example>
/// <example>
///   <summary>Resume a checkpoint job and retry recorded failures.</summary>
///   <code>Invoke-OfficeConversionJob -Path ./monthly.officejob.json -RetryFailed -ErrorAction Stop</code>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeConversionJob", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeBatchSummary), typeof(OfficeBatchItemReport), typeof(OfficeBatchPlan))]
public sealed class InvokeOfficeConversionJobCommand : AsyncPSCmdlet {
    /// <summary>Literal saved-job JSON path. Relative paths in the job resolve against this file's folder.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("LiteralPath", "FullName")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Inspect the job without converting files or creating output/checkpoint state.</summary>
    [Parameter]
    public SwitchParameter Plan { get; set; }

    /// <summary>Retry failed checkpoint items while verifying completed artifacts.</summary>
    [Parameter]
    public SwitchParameter RetryFailed { get; set; }

    /// <summary>Stream per-file outcomes instead of only the final bounded summary.</summary>
    [Parameter]
    public SwitchParameter ItemResults { get; set; }

    /// <summary>Variable receiving the completed plan or batch summary, including when ItemResults is selected.</summary>
    [Parameter]
    public string? SummaryVariable { get; set; }

    /// <summary>Allow execution to select no supported inputs without reporting an error.</summary>
    [Parameter]
    public SwitchParameter AllowEmpty { get; set; }

    /// <summary>New run-report JSON path outside the job's source, output and checkpoint trees. Existing files are preserved.</summary>
    [Parameter]
    public string? ReportPath { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        if (ReportPath is not null && string.IsNullOrWhiteSpace(ReportPath))
            throw new ArgumentException("Specify a non-empty run-report path.", nameof(ReportPath));
        string jobPath = OfficeFileCommandUtilities.ResolvePath(this, Path);
        string baseDirectory = System.IO.Path.GetDirectoryName(jobPath)!;
        OfficeBatchOptions job = OfficeBatchOptions.Load(jobPath);
        if (RetryFailed.IsPresent) job.RetryFailed = true;
        OfficeWorkflowCommandUtilities.RegisterConversions();
        if (Plan.IsPresent) {
            if (!string.IsNullOrWhiteSpace(ReportPath))
                throw new ArgumentException("Plan does not write a run report. Use ReportPath when executing the job.");
            var plan = await OfficeConverter.PlanAsync(job, cancellationToken: CancelToken);
            SetSummary(plan);
            WriteObject(plan);
            if (!plan.CanRun && !(AllowEmpty.IsPresent && plan.Selected == 0 && plan.Failed == 0))
                WriteError(new ErrorRecord(new InvalidOperationException("The conversion job plan requires attention. Inspect the plan diagnostics before execution."),
                    "OfficeConversionJobPlanInvalid", ErrorCategory.InvalidResult, plan));
            return;
        }
        string outputDirectory = job.OutputDirectory;
        if (!ShouldProcess(outputDirectory, "Run saved conversion job")) return;
        job.ResolveEnvironmentVariable = Environment.GetEnvironmentVariable;
        job.AllowEmpty |= AllowEmpty.IsPresent;
        job.ReportPath = string.IsNullOrWhiteSpace(ReportPath) ? null : OfficeFileCommandUtilities.ResolvePath(this, ReportPath);
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(job, new BatchProgress(this), CancelToken);
        OfficeBatchSummary result = report.Summary;
        WriteProgress(new ProgressRecord(1, "Run conversion job", "Finished") { RecordType = ProgressRecordType.Completed });
        OfficeWorkflowCommandUtilities.WriteSavedJobOutcome(this, result, report, SummaryVariable, AllowEmpty.IsPresent, ItemResults.IsPresent);
    }

    private void SetSummary(object result) {
        if (!string.IsNullOrWhiteSpace(SummaryVariable)) SessionState.PSVariable.Set(SummaryVariable, result);
    }

    private sealed class BatchProgress(InvokeOfficeConversionJobCommand command) : IProgress<OfficeBatchItemReport> {
        private int _processed;
        public void Report(OfficeBatchItemReport item) {
            if (command.ItemResults.IsPresent) command.WriteObject(item);
            int processed = System.Threading.Interlocked.Increment(ref _processed);
            if (processed == 1 || processed % 100 == 0)
                command.WriteProgress(new ProgressRecord(1, "Run conversion job", $"Processed {processed} files"));
        }
    }
}
