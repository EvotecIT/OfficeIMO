using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class StudioJobRecord {
    private const int MaximumDisplayedItems = 50;
    internal OfficeConversionJobRunReport? RunReport { get; private set; }
    internal string? RunReportPath { get; private set; }
    public bool HasRunReport => RunReport is not null;
    public bool CanReviewJob => HasRunReport && !IsActive;
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

    internal void ApplyRunReport(OfficeConversionJobRunReport report, string? reportPath) {
        // Keep only metadata and totals on the card; the file owns the complete bounded report.
        RunReport = report with { Batch = report.Batch is { } observation ? observation with { Items = [] } : null };
        RunReportPath = reportPath;
        Started = report.StartedUtc.ToLocalTime();
        FinishedLabel = report.FinishedUtc is { } finished
            ? _localizer.FormatOrDefault("Jobs.RunReport.Finished", "Finished {0}", finished.ToLocalTime().ToString("g")) : null;
        RecordingError = report.RecordingError is { } error
            ? _localizer.FormatOrDefault("Jobs.RunReport.RecordingError", "Run history could not be fully saved: {0}", error) : null;
        ReportItems = (report.Batch?.Items ?? []).Take(MaximumDisplayedItems).Select(item => new StudioConversionRunItem(item, _localizer)).ToArray();
        ReportSample = report.Batch is { } batch && (batch.ItemsTruncated || batch.Items.Count > MaximumDisplayedItems)
            ? _localizer.FormatOrDefault("Jobs.RunReport.Sample", "Showing {0} file results. The counts describe the whole attempt.", ReportItems.Count)
            : ReportItems.Count == 0 ? _localizer.GetOrDefault("Jobs.RunReport.NoFileResults", "No file results were recorded.") : null;
        ReportNote = _localizer.GetOrDefault("Jobs.RunReport.Note", "Saved attempts describe what happened. Checkpointed runs verify outputs again before reuse.");
        if (report.State == OfficeConversionJobRunState.Running) {
            _cancel = null;
            Status = _localizer.GetOrDefault("Jobs.Running", "Running");
            Summary = _localizer.GetOrDefault("Jobs.RunReport.RunningNote", "This saved job is running in another Studio instance. Review it after that attempt finishes.");
            CancelCommand.NotifyCanExecuteChanged();
        } else {
            OfficeWorkflowStatus outcome = report.State switch {
                OfficeConversionJobRunState.Completed => OfficeWorkflowStatus.Completed,
                OfficeConversionJobRunState.Cancelled => OfficeWorkflowStatus.Cancelled,
                OfficeConversionJobRunState.Failed when report.Batch?.Items.Any(item => item.Status == OfficeWorkflowStatus.Unconfirmed) == true => OfficeWorkflowStatus.Unconfirmed,
                _ => OfficeWorkflowStatus.Failed
            };
            var summary = report.Batch?.Summary;
            string text = summary is null ? report.FailureSummary ?? string.Empty
                : _localizer.FormatOrDefault("Jobs.RunReport.Summary", "{0} completed · {1} reused · {2} failed · {3} skipped",
                    summary.Completed, summary.Reused, summary.Failed, summary.Skipped);
            if (summary?.FailureSummary is { } failure) text += Environment.NewLine + failure;
            if (report.FailureSummary is { } runFailure && runFailure != summary?.FailureSummary) text += Environment.NewLine + runFailure;
            bool recordedOutput = summary is not null && (summary.Completed > 0 || summary.Reused > 0);
            bool availableOutput = recordedOutput && Directory.Exists(report.OutputDirectory);
            CompleteBatch(outcome, recordedOutput ? report.OutputDirectory : null, text, [], availableOutput, isDirectoryOutput: true);
            if (report.State == OfficeConversionJobRunState.Interrupted) {
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
    internal StudioConversionRunItem(OfficeConversionBatchItemResult item, IStudioLocalizer localizer) {
        InputPath = item.InputPath; InputName = Path.GetFileName(item.InputPath);
        OutputPath = item.OutputPath;
        Status = item.Reused ? localizer.GetOrDefault("Jobs.RunReport.Reused", "Reused")
            : item.Skipped ? localizer.GetOrDefault("Jobs.RunReport.Skipped", "Skipped")
            : item.Status is { } status ? localizer.GetOrDefault("Workflow.Status." + status, status.ToString()) : string.Empty;
        Summary = Bound(item.Summary) + string.Concat(item.Diagnostics
            .Where(diagnostic => diagnostic.Severity is OfficeWorkflowDiagnosticSeverity.Warning or OfficeWorkflowDiagnosticSeverity.Error)
            .Take(3).Select(diagnostic => Environment.NewLine + Bound(diagnostic.Message)));
        HasOutput = item.Status == OfficeWorkflowStatus.Completed && !string.IsNullOrWhiteSpace(OutputPath);
        IsFailed = item.Status is OfficeWorkflowStatus.Failed or OfficeWorkflowStatus.Unconfirmed;
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
