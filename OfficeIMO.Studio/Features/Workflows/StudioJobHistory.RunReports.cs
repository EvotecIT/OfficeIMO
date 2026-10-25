using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class StudioJobHistory {
    internal StudioConversionRunStore? RunStore { get; private set; }
    internal string MissingOutputMessage => _localizer.GetOrDefault("Jobs.RunReport.MissingOutput", "The output is no longer available.");

    [ObservableProperty] private string? _runHistoryError;

    private void InitializeRunReports(StudioDocumentHistory? documentHistory) {
        if (documentHistory is null) return;
        RunStore = documentHistory.ConversionRuns;
        var history = RunStore.ReadLatest();
        if (history.Error is { } error) RunHistoryError = _localizer.FormatOrDefault("Jobs.RunReport.LoadError", "Some saved run reports could not be read: {0}", error);
        foreach (var saved in history.Entries.Reverse()) {
            var report = saved.Report;
            var entry = Start(_localizer.GetOrDefault("Jobs.RunReport.Title", "Saved folder conversion"),
                report.JobPath, report.OutputDirectory, () => { }, batch: true);
            entry.ApplyRunReport(report, saved.ReportPath);
        }
        documentHistory.Cleared += (_, _) => {
            foreach (var entry in _entries.Where(entry => entry.HasRunReport && !entry.IsActive && !entry.HasRecovery &&
                         entry.RunReportPath is not null && !File.Exists(entry.RunReportPath)).ToArray()) Remove(entry);
            NotifyState();
        };
    }

    internal void RecordRun(StudioJobRecord? entry, OfficeConversionJobRunReport report, string reportPath) {
        entry?.ApplyRunReport(report, File.Exists(reportPath) ? reportPath : null);
        if (RunStore?.TrimFinished() == false)
            RunHistoryError = _localizer.GetOrDefault("Jobs.RunReport.ClearError", "Some saved run reports could not be cleared. Active attempts are retained.");
        NotifyState();
    }
}
