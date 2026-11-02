using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class StudioJobRecord {
    private const int MaximumDisplayedItems = 50;
    internal OfficeBatchReport? RunReport { get; private set; }
    internal string? RunReportPath { get; private set; }
    public bool HasRunReport => RunReport is not null;
    public bool CanReviewJob => RunReport is { JobPath: not null, JobSha256: not null } && !IsActive && !IsExternalRunning;
    internal bool IsObservedRun { get; private set; }

    /// <summary>A live report lease belongs to another instance; this window neither waits for nor cancels it.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanReviewJob))]
    private bool _isExternalRunning;
    public IReadOnlyList<StudioConversionRunItem> ReportItems { get; private set; } = [];
    public string? ReportNote { get; private set; }
    public string? RecordingError { get; private set; }
    public string? ReportSample { get; private set; }
    public string? OutputAvailabilityNote { get; private set; }
    public string? FinishedLabel { get; private set; }

    internal void SetRecordingError(string error) {
        RecordingError = _localizer.FormatOrDefault("Jobs.RunReport.RecordingError", "Run history could not be fully saved: {0}", error);
        OnPropertyChanged(nameof(RecordingError));
    }

    internal void OwnRunReportPath(string reportPath) => RunReportPath = reportPath;

    internal void ApplyRunReport(OfficeBatchReport report, string? reportPath, bool observed = false) {
        IsObservedRun |= observed;
        // Keep only metadata and totals on the card; the file owns the complete bounded report.
        RunReport = report with { Items = [] };
        RunReportPath = reportPath;
        Started = report.StartedUtc.ToLocalTime();
        FinishedLabel = report.FinishedUtc is { } finished
            ? _localizer.FormatOrDefault("Jobs.RunReport.Finished", "Finished {0}", finished.ToLocalTime().ToString("g")) : null;
        RecordingError = report.RecordingError is { } error
            ? _localizer.FormatOrDefault("Jobs.RunReport.RecordingError", "Run history could not be fully saved: {0}", error) : null;
        ReportItems = report.Items.Take(MaximumDisplayedItems).Select(item => new StudioConversionRunItem(item, _localizer)).ToArray();
        ReportSample = report is { } batch && (batch.ItemsTruncated || batch.Items.Count > MaximumDisplayedItems)
            ? _localizer.FormatOrDefault("Jobs.RunReport.Sample", "Showing {0} file results. The counts describe the whole attempt.", ReportItems.Count)
            : ReportItems.Count == 0 ? _localizer.GetOrDefault("Jobs.RunReport.NoFileResults", "No file results were recorded.") : null;
        ReportNote = _localizer.GetOrDefault("Jobs.RunReport.Note", "Saved attempts describe what happened. Checkpointed runs verify outputs again before reuse.");
        IsExternalRunning = observed && reportPath is not null && OfficeBatchReport.IsActive(reportPath);
        if (IsExternalRunning || report.State == OfficeBatchRunState.Running) {
            _cancel = null;
            IsActive = false;
            Outcome = null; Progress = 0; HasOutput = false; OutputPath = null;
            Status = _localizer.GetOrDefault("Jobs.RunReport.RunningElsewhere", "Running elsewhere");
            Summary = _localizer.GetOrDefault("Jobs.RunReport.RunningNote", "This saved job is running in another Studio instance. Review it after that attempt finishes.");
            CancelCommand.NotifyCanExecuteChanged();
        } else {
            OfficeWorkflowStatus outcome = report.State switch {
                OfficeBatchRunState.Completed => OfficeWorkflowStatus.Completed,
                OfficeBatchRunState.Cancelled => OfficeWorkflowStatus.Cancelled,
                OfficeBatchRunState.Failed when report.Items.Any(item => item.Status == OfficeBatchItemStatus.Unconfirmed) => OfficeWorkflowStatus.Unconfirmed,
                _ => OfficeWorkflowStatus.Failed
            };
            var summary = report.Summary;
            string text = summary is null ? report.FailureSummary ?? string.Empty
                : _localizer.FormatOrDefault("Jobs.RunReport.Summary", "{0} completed · {1} reused · {2} failed · {3} skipped",
                    summary.Completed, summary.Reused, summary.Failed, summary.Skipped);
            if (summary?.FailureSummary is { } failure) text += Environment.NewLine + failure;
            if (report.FailureSummary is { } runFailure && runFailure != summary?.FailureSummary) text += Environment.NewLine + runFailure;
            bool recordedOutput = summary is not null && (summary.Completed > 0 || summary.Reused > 0);
            bool availableOutput = recordedOutput && Directory.Exists(report.OutputDirectory);
            ApplyBatchCompletion(outcome, recordedOutput ? report.OutputDirectory : null, text, [], availableOutput, isDirectoryOutput: true);
            if (report.State == OfficeBatchRunState.Interrupted) {
                Status = _localizer.GetOrDefault("Jobs.RunReport.Interrupted", "Interrupted");
                Summary = _localizer.GetOrDefault("Jobs.RunReport.InterruptedNote", "This attempt ended without a final result. Review the saved job and outputs before running again.");
            }
            OutputAvailabilityNote = recordedOutput && !availableOutput
                ? _localizer.GetOrDefault("Jobs.RunReport.OutputUnavailable", "Recorded output is no longer available.") : null;
        }
        OnPropertyChanged(nameof(StartedLabel)); OnPropertyChanged(nameof(FinishedLabel));
        OnPropertyChanged(nameof(HasRunReport)); OnPropertyChanged(nameof(CanReviewJob)); OnPropertyChanged(nameof(CanCancel));
        OnPropertyChanged(nameof(ReportItems)); OnPropertyChanged(nameof(ReportNote)); OnPropertyChanged(nameof(RecordingError));
        OnPropertyChanged(nameof(ReportSample)); OnPropertyChanged(nameof(OutputAvailabilityNote));
    }
}

/// <summary>Bounded presentation of one observed file result; opening rechecks its current location.</summary>
public sealed class StudioConversionRunItem {
    internal StudioConversionRunItem(OfficeBatchItemReport item, IStudioLocalizer localizer) {
        InputPath = item.InputPath; InputName = Path.GetFileName(item.InputPath);
        OutputPath = item.OutputPath;
        Status = item.Reused ? localizer.GetOrDefault("Jobs.RunReport.Reused", "Reused")
            : item.Skipped ? localizer.GetOrDefault("Jobs.RunReport.Skipped", "Skipped")
            : item.Status is { } status ? localizer.GetOrDefault("Workflow.Status." + status, status.ToString()) : string.Empty;
        Summary = Bound(item.Summary) + string.Concat(item.Diagnostics
            .Where(diagnostic => diagnostic.Severity is OfficeDiagnosticSeverity.Warning or OfficeDiagnosticSeverity.Error)
            .Take(3).Select(diagnostic => Environment.NewLine + Bound(diagnostic.Message)));
        HasOutput = item.Status == OfficeBatchItemStatus.Completed && !string.IsNullOrWhiteSpace(OutputPath);
        IsFailed = item.Status is OfficeBatchItemStatus.Failed or OfficeBatchItemStatus.Unconfirmed;
    }
    public string InputName { get; }
    public string InputPath { get; }
    public string? OutputPath { get; }
    public string Status { get; }
    public string Summary { get; }
    public bool HasOutput { get; }
    public bool IsFailed { get; }
    private static string Bound(string value) => value.Length > 1024 ? value[..1024] + "…" : value;
}
