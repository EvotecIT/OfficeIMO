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
        RefreshRunReports();
        documentHistory.Cleared += (_, _) => {
            foreach (var entry in _entries.Where(entry => entry.HasRunReport && !entry.IsActive && !entry.HasRecovery &&
                         entry.RunReportPath is not null && !File.Exists(entry.RunReportPath)).ToArray()) Remove(entry);
            NotifyState();
        };
    }

    /// <summary>Reconciles a bounded snapshot in place without taking ownership of another instance's work.</summary>
    internal void RefreshRunReports() {
        if (RunStore is null) return;
        var history = RunStore.ReadLatest();
        string? error = history.Error;
        var reports = history.Entries.ToList();
        var comparer = OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
        if (!RunStore.CanRetainReport) {
            // Opt-out stops new admissions. Existing visible history can still observe its writer finishing.
            foreach (var entry in _entries.Where(entry => entry.IsObservedRun && entry.RunReportPath is not null)
                         .Take(StudioConversionRunStore.MaximumEntries)) {
                try {
                    if (File.Exists(entry.RunReportPath)) reports.Add(new(entry.RunReportPath!, RunStore.ReadReport(entry.RunReportPath!)));
                } catch (Exception failure) when (failure is not OutOfMemoryException and not StackOverflowException) { error ??= failure.Message; }
            }
        }
        var paths = reports.Select(saved => saved.ReportPath).ToHashSet(comparer);
        foreach (var saved in reports.AsEnumerable().Reverse()) {
            try {
                StudioJobRecord? existing = _entries.FirstOrDefault(entry => entry.RunReportPath is { } path && comparer.Equals(path, saved.ReportPath));
                if (existing is not null) {
                    if (existing.IsObservedRun && !existing.IsActive) existing.ApplyRunReport(saved.Report, saved.ReportPath, observed: true);
                    continue;
                }
                var report = saved.Report;
                var entry = new StudioJobRecord(_localizer.GetOrDefault("Jobs.RunReport.Title", "Saved folder conversion"),
                    report.JobPath, report.OutputDirectory, null, true, _localizer);
                entry.ApplyRunReport(report, saved.ReportPath, observed: true);
                AddEntry(entry);
            } catch (Exception failure) when (failure is not OutOfMemoryException and not StackOverflowException) { error ??= failure.Message; }
        }
        foreach (var entry in _entries.Where(entry => entry.IsObservedRun && !entry.IsActive && !entry.HasRecovery &&
                     entry.RunReportPath is not null && !paths.Contains(entry.RunReportPath)).ToArray()) {
            try {
                bool missing = !File.Exists(entry.RunReportPath);
                if (missing && OfficeConversionJobRunReportSerializer.IsActive(entry.RunReportPath!)) continue;
                if (missing || RunStore.CanRetainReport && history.Error is null) Remove(entry);
            } catch (Exception failure) when (failure is not OutOfMemoryException and not StackOverflowException) { error ??= failure.Message; }
        }
        RunHistoryError = error is null ? null : _localizer.FormatOrDefault("Jobs.RunReport.LoadError",
            "Some saved run reports could not be read: {0}", error.Length > 1024 ? error[..1024] : error);
        NotifyState();
    }

    internal void RecordRun(StudioJobRecord? entry, OfficeConversionJobRunReport report, string reportPath) {
        entry?.ApplyRunReport(report, File.Exists(reportPath) ? reportPath : null);
        if (RunStore?.TrimFinished() == false)
            RunHistoryError = _localizer.GetOrDefault("Jobs.RunReport.ClearError", "Some saved run reports could not be cleared. Active attempts are retained.");
        NotifyState();
    }
}
