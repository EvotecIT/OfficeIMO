using System.ComponentModel;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class SearchablePdfOcrViewModel {
    private void ScanChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName == nameof(ScanPreparationViewModel.Dpi)) RenderDpi = Scan.Dpi;
        if (e.PropertyName == nameof(ScanPreparationViewModel.IsBusy)) {
            RunCommand.NotifyCanExecuteChanged(); ExtractTextCommand.NotifyCanExecuteChanged();
        }
        if (e.PropertyName == nameof(ScanPreparationViewModel.PageNumber)) ExtractedText = string.Empty;
    }
    partial void OnRenderDpiChanged(double value) { if (Scan != null) Scan.Dpi = value; }
    private async Task<byte[]> ReadScanSourceAsync(CancellationToken token) {
        if (string.IsNullOrWhiteSpace(InputPath)) throw new InvalidOperationException("Choose a PDF before previewing its scans.");
        if (_storage != null) return (await _storage.ReadSnapshotAsync(InputPath, token, 128L * 1024 * 1024).ConfigureAwait(true)).Bytes;
        using var access = new StudioStorageAccess();
        return (await access.ReadSnapshotAsync(InputPath, token, 128L * 1024 * 1024).ConfigureAwait(true)).Bytes;
    }
    private async Task SaveScanCopyAsync(PdfOcrMergeOptions options, string sourceHash, CancellationToken token) {
        if (IsBusy) return;
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(token);
        _cancellation = operation; IsBusy = true;
        string input = InputPath;
        StudioJobRecord? job = null;
        try {
            string? output = await _pickOutputPdf(operation.Token).ConfigureAwait(true);
            if (string.IsNullOrWhiteSpace(output)) return;
            bool provider = _storage?.UsesProviderPublication(output) == true;
            if (!provider && !_canPublishPath(output))
                throw new InvalidOperationException(T("Error.OutputOpen", "That PDF is already open in another tab. Close it or choose a different output file name."));
            if (provider && !await _confirmProviderWrite(output).ConfigureAwait(true)) return;
            operation.Token.ThrowIfCancellationRequested();
            PublishedPath = null; ErrorMessage = null; HasRecovery = false;
            job = _jobHistory?.Start(_localizer.GetOrDefault("Scan.Title", "Scan preparation"), input, output, operation.Cancel);
            using IDisposable? execution = _jobHistory == null ? null : await _jobHistory.EnterAsync(operation.Token).ConfigureAwait(true);
            var request = new OfficeWorkflowRequest {
                Operation = OfficeWorkflowOperation.ScanCleanup,
                InputPath = input,
                InputStream = _storage?.CreateWorkflowInput(input),
                OutputPath = output,
                PublicationGuard = _publicationGuard,
                ConflictPolicy = provider || ReplaceExistingOutput ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail,
                OutputStream = provider ? _storage!.CreateWorkflowOutput(output, _recoveryStore ?? throw new IOException("Workflow recovery storage is unavailable.")) : null,
                ScanCleanup = new() { Preparation = options, ExpectedSourceSha256 = sourceHash, AcknowledgeRasterOutput = true }
            };
            var result = await new OfficeWorkflowRunner().RunAsync(request, cancellationToken: operation.Token).ConfigureAwait(true);
            Status = result.Summary; Summary = result.Summary;
            HasRecovery = result.Recovery != null;
            job?.Complete(result.Status, result.OutputPath, result.Summary, result.Recovery);
            if (result.Status == OfficeWorkflowStatus.Completed) PublishedPath = result.OutputPath;
            else ErrorMessage = result.Summary;
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            job?.Complete(OfficeWorkflowStatus.Cancelled, null, T("Status.Cancelled", "OCR cancelled"));
            throw;
        } catch (Exception error) {
            job?.Complete(OfficeWorkflowStatus.Failed, null, error.Message);
            throw;
        } finally { if (ReferenceEquals(_cancellation, operation)) _cancellation = null; IsBusy = false; }
    }
    partial void OnIsBusyChanged(bool value) { if (Scan != null) Scan.HostBusy = value; }
}
