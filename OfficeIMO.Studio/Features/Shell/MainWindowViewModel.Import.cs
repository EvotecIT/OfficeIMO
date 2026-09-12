using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private readonly Func<PageImportPreviewViewModel, Task<bool>> _reviewPageImport;
    private bool _reviewingPageImport;

    [RelayCommand]
    private async Task ImportPagesAsync(CancellationToken cancellationToken) {
        using var notifications = BeginNotificationScope();
        if (_workspace is null || !CanImportPages || IsWorkspaceBusy || _reviewingPageImport) return;
        PdfWorkspace workspace = _workspace;
        long revision = workspace.Revision;
        int insertBefore = _organizerSelection.Count == 0 ? workspace.Pages.Count + 1 : _organizerSelection.Min();
        _reviewingPageImport = true;
        try {
            IReadOnlyList<string> paths = await _pickImportPdfs(cancellationToken).ConfigureAwait(true);
            if (paths.Count == 0 || !IsPageWorkflowCurrent(workspace, revision) || !CanImportPages) return;
            PdfImportPreparation? preparation = null;
            bool prepared = await RunStandaloneAsync(async token => preparation = await workspace
                .PrepareImportAsync(paths, token, _promptPdfPassword).ConfigureAwait(true), cancellationToken).ConfigureAwait(true);
            if (!prepared || !IsPageWorkflowCurrent(workspace, revision)) return;
            if (preparation is null) { OperationStatus = UiText("Workspace.OperationCancelled"); return; }
            var preview = new PageImportPreviewViewModel(preparation, insertBefore, _localizer);
            if (!await _reviewPageImport(preview).ConfigureAwait(true) || !preview.CanApply) return;
            if (!IsPageWorkflowCurrent(workspace, revision) || !CanImportPages) return;
            int position = preview.InsertBefore;
            PdfImportSelection[] selections = preview.Selections;
            int readingPage = SelectedPage?.PageNumber ?? 1;
            int importedCount = 0;
            bool succeeded = await RunStandaloneAsync(async token => {
                var job = _services.Jobs.Start(UiText("Organizer.ImportTitle"), workspace.Path, null, CancelCurrentOperation);
                try {
                    using IDisposable execution = await _services.Jobs.EnterAsync(token).ConfigureAwait(true);
                    var progress = new Progress<PdfWorkspaceProgress>(update => {
                        if (!job.IsActive) return;
                        OperationStatus = update.Stage; OperationProgressFraction = update.Fraction;
                        job.Report(new OfficeWorkflowProgress("import", "import", update.Stage, update.Fraction, update.Fraction));
                    });
                    importedCount = await workspace.ApplyImportAsync(preparation, selections, position, token, progress).ConfigureAwait(true);
                    job.Complete(OfficeWorkflowStatus.Completed, null, UiFormat("Workspace.ImportedPages", importedCount, selections.Length));
                } catch (OperationCanceledException) when (token.IsCancellationRequested) {
                    job.Complete(OfficeWorkflowStatus.Cancelled, null, UiText("Workspace.OperationCancelled")); throw;
                } catch (Exception error) { job.Complete(OfficeWorkflowStatus.Failed, null, error.Message); throw; }
            }, cancellationToken).ConfigureAwait(true);
            if (succeeded && ReferenceEquals(workspace, _workspace) && !_disposed) {
                RefreshWorkspacePresentation(Enumerable.Range(position, importedCount).ToArray());
                NavigateToPage(readingPage >= position ? readingPage + importedCount : readingPage);
                OperationStatus = importedCount == 1 ? UiText("Workspace.ImportedOnePage") : UiFormat("Workspace.ImportedPages", importedCount, selections.Length);
            }
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { OperationStatus = UiText("Workspace.OperationCancelled"); }
        catch (Exception error) { ErrorMessage = error.Message; }
        finally { _reviewingPageImport = false; }
    }
}
