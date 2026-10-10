using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    [RelayCommand]
    private async Task ExportComparisonReportAsync(CancellationToken cancellationToken) {
        if (!CanExportComparisonReport || _workspace is not { } workspace || _comparisonReport is not { } report ||
            _session is not { } primary || _comparisonSession is not { } actual) return;
        long revision = _comparisonReportRevision;
        bool IsCurrent() => !_disposed && !HasFormDrafts && ReferenceEquals(_workspace, workspace) && ReferenceEquals(_comparisonReport, report) &&
            ReferenceEquals(_session, primary) && ReferenceEquals(_comparisonSession, actual) && workspace.Revision == revision;
        await RunStandaloneAsync(async token => {
            string? destination = await _pickSaveComparisonReport(token).ConfigureAwait(true);
            if (destination is null || !IsCurrent()) return;
            await workspace.ExportComparisonReportAsync(destination, report, revision, actual.Path, token, IsCurrent).ConfigureAwait(true);
            if (IsCurrent()) OperationStatus = ComparisonText("ReportSaved", "Saved the standalone comparison report. Both source PDFs are unchanged.");
        }, cancellationToken).ConfigureAwait(true);
    }
}
