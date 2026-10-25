using System;
using System.IO;
using System.Management.Automation;
using System.Threading.Tasks;
using PSWriteOffice.Services;
#if !FRAMEWORK
using OfficeIMO.Workflows;
#endif

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Plans or runs a saved conversion job through the same runner as Studio and the CLI.</summary>
/// <para>Requires PowerShell 7.4 or newer. Plan performs no output or checkpoint writes. Execution revalidates current inputs and retains completed artifacts after a partial failure or cancellation. Password environment references are resolved only for execution; values are not written to the job or reports.</para>
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
#if !FRAMEWORK
[OutputType(typeof(OfficeConversionBatchResult), typeof(OfficeConversionBatchItemResult), typeof(OfficeConversionJobPlan))]
#endif
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

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
#if FRAMEWORK
        await Task.CompletedTask;
        throw new PlatformNotSupportedException("Saved conversion jobs require PowerShell 7.4 or newer.");
#else
        string jobPath = OfficeFileCommandUtilities.ResolvePath(this, Path);
        string baseDirectory = System.IO.Path.GetDirectoryName(jobPath)!;
        OfficeConversionJob job = OfficeConversionJobSerializer.ReadFile(jobPath);
        if (RetryFailed.IsPresent) job.RetryFailed = true;
        var runner = OfficeWorkflowCommandUtilities.CreateRunner();
        if (Plan.IsPresent) {
            var plan = await OfficeConversionJobRunner.PlanAsync(job, baseDirectory, runner, cancellationToken: CancelToken);
            SetSummary(plan);
            WriteObject(plan);
            if (!plan.CanRun && !(AllowEmpty.IsPresent && plan.Selected == 0 && plan.Failed == 0))
                WriteError(new ErrorRecord(new InvalidOperationException("The conversion job plan requires attention. Inspect the plan diagnostics before execution."),
                    "OfficeConversionJobPlanInvalid", ErrorCategory.InvalidResult, plan));
            return;
        }
        string outputDirectory = System.IO.Path.GetFullPath(job.OutputDirectory, baseDirectory);
        if (!ShouldProcess(outputDirectory, "Run saved conversion job")) return;
        var result = await OfficeConversionJobRunner.RunAsync(job, baseDirectory, runner,
            new OfficeConversionJobRuntimeOptions { ResolveEnvironmentVariable = Environment.GetEnvironmentVariable },
            new BatchProgress(this), CancelToken);
        SetSummary(result);
        WriteProgress(new ProgressRecord(1, "Run conversion job", "Finished") { RecordType = ProgressRecordType.Completed });
        ErrorRecord? error = OfficeWorkflowCommandUtilities.GetBatchError(result, AllowEmpty.IsPresent, "OfficeConversionJob");
        if (error != null) WriteError(error);
        if (!ItemResults.IsPresent) WriteObject(result);
#endif
    }

    private void SetSummary(object result) {
        if (!string.IsNullOrWhiteSpace(SummaryVariable)) SessionState.PSVariable.Set(SummaryVariable, result);
    }

#if !FRAMEWORK
    private sealed class BatchProgress(InvokeOfficeConversionJobCommand command) : IProgress<OfficeConversionBatchItemResult> {
        private int _processed;
        public void Report(OfficeConversionBatchItemResult item) {
            if (command.ItemResults.IsPresent) command.WriteObject(item);
            int processed = System.Threading.Interlocked.Increment(ref _processed);
            if (processed == 1 || processed % 100 == 0)
                command.WriteProgress(new ProgressRecord(1, "Run conversion job", $"Processed {processed} files"));
        }
    }
#endif
}
