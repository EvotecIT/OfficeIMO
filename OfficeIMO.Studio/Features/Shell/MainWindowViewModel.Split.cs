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
            bool provider = _services.Storage.UsesProviderPublication(folder);
            string destination = provider ? folder : Path.Combine(OfficeStorageIdentity.GetLocalPath(folder)
                ?? throw new IOException("Choose an accessible output folder."), "Split PDFs");
            var preview = new PageSplitPreviewViewModel(workspace.Pages.Count, pagesPerPart, destination, provider, _localizer,
                path => _openDocumentInTab is null ? _openUri(new Uri(path)) : _openDocumentInTab(path, CancellationToken.None));
            if (!await _reviewPageSplit(preview).ConfigureAwait(true) || !preview.CanApply) return;
            if (!IsReviewedCopyCurrent(workspace, revision) || !CanExtractPages) return;
            if (provider && !await _confirmProviderWrite(destination).ConfigureAwait(true)) return;
            if (!IsReviewedCopyCurrent(workspace, revision) || !CanExtractPages) return;
            int selectedPartSize = preview.PartSize;
            SplitPagesPerDocument = selectedPartSize;
            PdfSplitWorkflowResult? result = null;
            await RunStandaloneAsync(async token => {
                StudioJobRecord? job = null;
                StudioStorageAccess.DirectoryOutputSession? directory = null;
                bool ownerStarted = false;
                try {
                    directory = provider ? _services.Storage.CreateDirectoryOutput(destination, _services.WorkflowRecovery) : null;
                    job = _services.Jobs.Start(UiText("Organizer.SplitTitle"), workspace.Path, destination, CancelCurrentOperation);
                    using IDisposable execution = await _services.Jobs.EnterAsync(token).ConfigureAwait(true);
                    if (!IsReviewedCopyCurrent(workspace, revision)) throw new InvalidOperationException(ErrorMessage);
                    var progress = new Progress<Features.Workspace.PdfWorkspaceProgress>(update => {
                        if (!job.IsActive) return;
                        OperationStatus = update.Stage; OperationProgressFraction = update.Fraction;
                        job.Report(new OfficeWorkflowProgress("split", "split", update.Stage, update.Fraction, update.Fraction));
                    });
                    ownerStarted = true;
                    result = await workspace.SplitAsync(destination, selectedPartSize, token, progress, directory?.Output, _publicationGuard).ConfigureAwait(true);
                    string? publishedOutput = result.Files.Count == 0 ? null : provider ? result.Files[0].Path : Path.GetDirectoryName(result.Files[0].Path);
                    job.CompleteBatch(result.Status, publishedOutput, result.Summary, result.OutputRecoveries, result.Files.Count > 0);
                } catch (OperationCanceledException) when (token.IsCancellationRequested) {
                    // The runner returns typed cancellation/uncertainty after starting; an escaping cancellation
                    // comes from admission to the jobs or workspace CPU gate before output work starts.
                    job?.Complete(OfficeWorkflowStatus.Cancelled, null, UiText("Workspace.OperationCancelled"));
                    throw;
                } catch (Exception error) {
                    if (ownerStarted) job?.Unconfirmed(error.Message);
                    else job?.Complete(OfficeWorkflowStatus.Failed, null, error.Message);
                    throw;
                } finally { directory?.Dispose(); }
            }, cancellationToken).ConfigureAwait(true);
            if (result is not null && !_disposed && ReferenceEquals(workspace, _workspace)) {
                OperationStatus = result.Summary;
                preview.Complete(result);
                await _showPageSplitResult(preview).ConfigureAwait(true);
            }
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            OperationStatus = UiText("Workspace.OperationCancelled");
        } catch (Exception error) { ErrorMessage = error.Message; }
        finally { _reviewingPageSplit = false; }
    }
}
