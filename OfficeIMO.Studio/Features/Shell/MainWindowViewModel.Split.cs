using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Internal;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private readonly Func<PageSplitPreviewViewModel, Task<bool>> _reviewPageSplit;
    private readonly Func<PageSplitPreviewViewModel, Task> _showPageSplitResult;
    private bool _reviewingPageSplit;

    [RelayCommand]
    private async Task SplitAsync(CancellationToken cancellationToken) {
        using var notifications = BeginNotificationScope();
        if (_workspace is null || !CanExtractPages || IsWorkspaceBusy || _reviewingPageSplit) return;
        var workspace = _workspace;
        long revision = workspace.Revision;
        if (!IsReviewedCopyCurrent(workspace, revision)) return;
        int pagesPerPart = SplitPagesPerDocument;
        _reviewingPageSplit = true;
        try {
            string? folder = await _pickOutputFolder(cancellationToken).ConfigureAwait(true);
            if (string.IsNullOrWhiteSpace(folder) || !IsReviewedCopyCurrent(workspace, revision)) return;
            string destination = Path.Combine(StudioStorageAccess.RequireLocalOutputFolder(folder, _localizer), "Split PDFs");
            PdfSplitStart[] starts = Bookmarks.Where(bookmark => bookmark.Level == 1 && bookmark.PageNumber.HasValue)
                .Select(bookmark => new PdfSplitStart(bookmark.PageNumber!.Value, bookmark.Title)).ToArray();
            var preview = new PageSplitPreviewViewModel(workspace.Pages.Count, pagesPerPart, destination, false, _localizer,
                path => _openDocumentInTab is null ? _openUri(new Uri(path)) : _openDocumentInTab(path, CancellationToken.None), starts);
            if (!await _reviewPageSplit(preview).ConfigureAwait(true) || !preview.CanApply) return;
            if (!IsReviewedCopyCurrent(workspace, revision) || !CanExtractPages) return;
            int selectedPartSize = preview.PartSize;
            PdfSplitPlan? selectedPlan = preview.Plan;
            SplitPagesPerDocument = selectedPartSize;
            PdfSplitWorkflowResult? result = null;
            await RunStandaloneAsync(async token => {
                StudioJobRecord? job = null;
                bool ownerStarted = false;
                try {
                    job = _services.Jobs.Start(UiText("Organizer.SplitTitle"), workspace.Path, destination, CancelCurrentOperation);
                    using IDisposable execution = await _services.Jobs.EnterAsync(token).ConfigureAwait(true);
                    if (!IsReviewedCopyCurrent(workspace, revision)) throw new InvalidOperationException(ErrorMessage);
                    var progress = new Progress<Features.Workspace.PdfWorkspaceProgress>(update => {
                        if (!job.IsActive) return;
                        OperationStatus = update.Stage; OperationProgressFraction = update.Fraction;
                        job.Report(new OfficeWorkflowProgress("split", "split", update.Stage, update.Fraction, update.Fraction));
                    });
                    ownerStarted = true;
                    result = await workspace.SplitAsync(destination, selectedPartSize, token, progress, null, _publicationGuard, selectedPlan).ConfigureAwait(true);
                    string? publishedOutput = result.Files.Count == 0 ? null : Path.GetDirectoryName(result.Files[0].Path);
                    job.CompleteBatch(result.Status, publishedOutput, result.Summary, result.OutputRecoveries, result.Files.Count > 0, isDirectoryOutput: true);
                } catch (OperationCanceledException) when (token.IsCancellationRequested) {
                    // The runner returns typed cancellation/uncertainty after starting; an escaping cancellation
                    // comes from admission to the jobs or workspace CPU gate before output work starts.
                    job?.Complete(OfficeWorkflowStatus.Cancelled, null, UiText("Workspace.OperationCancelled"));
                    throw;
                } catch (Exception error) {
                    if (ownerStarted) job?.Unconfirmed(error.Message);
                    else job?.Complete(OfficeWorkflowStatus.Failed, null, error.Message);
                    throw;
                }
            }, cancellationToken).ConfigureAwait(true);
            if (result is not null && !_disposed && ReferenceEquals(workspace, _workspace)) {
                OperationStatus = result.Summary;
                preview.Complete(result);
                await _showPageSplitResult(preview).ConfigureAwait(true);
            }
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            OperationStatus = UiText("Workspace.OperationCancelled");
        } catch (Exception error) { ErrorMessage = Infrastructure.StudioMessages.Describe(error); }
        finally { _reviewingPageSplit = false; }
    }
}
