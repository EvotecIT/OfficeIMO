using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private readonly Func<PageExtractionPreviewViewModel, Task<bool>> _reviewPageExtraction;
    private readonly Func<PageExtractionPreviewViewModel, Task> _showPageExtractionResult;
    private bool _reviewingPageExtraction;

    [RelayCommand]
    private async Task ExtractSelectedAsync(CancellationToken cancellationToken) {
        using var notifications = BeginNotificationScope();
        int[] pages = GetSelectedPages();
        if (_workspace is null || !CanExtractPages || pages.Length == 0 || IsWorkspaceBusy || _reviewingPageExtraction) return;
        var workspace = _workspace;
        long revision = workspace.Revision;
        if (!IsReviewedCopyCurrent(workspace, revision)) return;
        _reviewingPageExtraction = true;
        try {
            string? destination = await _pickSavePdf(cancellationToken).ConfigureAwait(true);
            if (string.IsNullOrWhiteSpace(destination) || !IsReviewedCopyCurrent(workspace, revision)) return;
            bool provider = _services.Storage.UsesProviderPublication(destination);
            var preview = new PageExtractionPreviewViewModel(workspace.Pages.Count, pages, destination, provider, _localizer,
                path => _openDocumentInTab is null ? _openUri(new Uri(path)) : _openDocumentInTab(path, CancellationToken.None),
                path => _openUri(new Uri(Path.GetDirectoryName(path)!)));
            if (!await _reviewPageExtraction(preview).ConfigureAwait(true) || !preview.CanApply) return;
            if (!IsReviewedCopyCurrent(workspace, revision) || !CanExtractPages) return;
            int[] selectedPages = preview.SelectedPages;
            if (provider && !await _confirmProviderWrite(destination).ConfigureAwait(true)) return;
            if (!IsReviewedCopyCurrent(workspace, revision) || !CanExtractPages) return;
            OfficeWorkflowResult? result = null;
            await RunStandaloneAsync(async token => {
                StudioJobRecord? job = null;
                try {
                    var output = _services.Storage.CreateWorkflowOutput(destination, _services.WorkflowRecovery);
                    job = _services.Jobs.Start(UiText("Organizer.ExtractTitle"), workspace.Path, destination, CancelCurrentOperation);
                    using IDisposable execution = await _services.Jobs.EnterAsync(token).ConfigureAwait(true);
                    if (!IsReviewedCopyCurrent(workspace, revision)) throw new InvalidOperationException(ErrorMessage);
                    var progress = new Progress<Features.Workspace.PdfWorkspaceProgress>(update => {
                        if (!job.IsActive) return;
                        OperationStatus = update.Stage; OperationProgressFraction = update.Fraction;
                        job.Report(new OfficeWorkflowProgress("extract", "extract", update.Stage, update.Fraction));
                    });
                    result = await workspace.ExtractAsync(selectedPages, destination, token, progress, output, _publicationGuard).ConfigureAwait(true);
                    job.Complete(result.Status, result.OutputPath, result.Summary, result.Recovery);
                } catch (OperationCanceledException) when (token.IsCancellationRequested) {
                    job?.Complete(OfficeWorkflowStatus.Cancelled, null, UiText("Workspace.OperationCancelled"));
                    throw;
                } catch (Exception error) {
                    // The shared runner returns publication failures and uncertainty as typed results.
                    // An ordinary escaping error comes from validation or admission before it starts.
                    job?.Complete(OfficeWorkflowStatus.Failed, null, error.Message);
                    throw;
                }
            }, cancellationToken).ConfigureAwait(true);
            if (result is not null && !_disposed && ReferenceEquals(workspace, _workspace)) {
                OperationStatus = result.Summary;
                if (result.Status is OfficeWorkflowStatus.Failed or OfficeWorkflowStatus.Unconfirmed) ErrorMessage = result.Summary;
                preview.Complete(result);
                await _showPageExtractionResult(preview).ConfigureAwait(true);
            }
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            OperationStatus = UiText("Workspace.OperationCancelled");
        } catch (Exception error) { ErrorMessage = error.Message; }
        finally { _reviewingPageExtraction = false; }
    }
}
