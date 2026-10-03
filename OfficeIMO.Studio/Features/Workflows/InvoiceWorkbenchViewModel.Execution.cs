using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class InvoiceWorkbenchViewModel {
    [RelayCommand(CanExecute = nameof(CanEdit))]
    private async Task ChooseInputAsync(CancellationToken token) {
        try {
            string? input = await _pickInvoice(token).ConfigureAwait(true);
            if (!_disposed && !string.IsNullOrWhiteSpace(input)) InputPath = input;
        } catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { Status = error.Message; }
    }
    [RelayCommand(CanExecute = nameof(CanEdit))]
    private async Task ChooseOutputFolderAsync(CancellationToken token) {
        try {
            string? folder = await _pickFolder(token).ConfigureAwait(true);
            if (!_disposed && !string.IsNullOrWhiteSpace(folder)) OutputFolder = folder;
        } catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { Status = error.Message; }
    }
    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync() {
        var cancellation = new CancellationTokenSource(); _cancellation = cancellation;
        IsBusy = true; OutputPath = null; Diagnostics.Clear();
        SourceSummary = ModelSummary = MappingSummary = SchemaSummary = RulesSummary = "—";
        Status = T("Running", "Processing invoice…");
        StudioJobRecord? job = null;
        StudioStorageAccess.DirectoryOutputSession? providerOutput = null;
        try {
            // Capture all UI choices before awaiting permission or a storage provider.
            var validator = CaptureValidator(); var request = CaptureRequest();
            string outputName = GetOutputName(), folder = OutputFolder, label = SelectedOperation.Label;
            if (request.OutputPath != null && !string.IsNullOrWhiteSpace(folder) && _storage?.UsesProviderPublication(folder) == true) {
                if (!await _confirmProviderWrite(folder).ConfigureAwait(true)) {
                    Status = T("BeforeWriteCancelled", "Cancelled before publishing."); return;
                }
                cancellation.Token.ThrowIfCancellationRequested();
                providerOutput = _storage.CreateDirectoryOutput(folder, _recovery ?? throw new IOException("Output recovery storage is unavailable."));
                var destination = await providerOutput!.ResolveAsync(outputName, cancellation.Token).ConfigureAwait(true);
                request = request with { OutputPath = destination.Location, OutputStream = destination.Output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace };
            }
            job = _jobs?.Start(label, request.InputPath, request.OutputPath, cancellation.Cancel);
            using IDisposable? execution = _jobs == null ? null : await _jobs.EnterAsync(cancellation.Token).ConfigureAwait(true);
            var result = await _runner.RunInvoiceAsync(request, validator, cancellation.Token).ConfigureAwait(true);
            job?.Complete(result.Status, result.OutputPath, result.Summary, result.Recovery);
            if (!_disposed) ShowReport(result, request.ValidationRelease.HasValue);
        } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
            Status = T("Cancelled", "Invoice operation cancelled."); job?.Complete(OfficeWorkflowStatus.Cancelled, null, Status);
        } catch (Exception error) {
            Status = error.Message; job?.Complete(OfficeWorkflowStatus.Failed, null, Status);
        } finally {
            try { providerOutput?.Dispose(); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) { Status += " " + error.Message; }
            if (ReferenceEquals(_cancellation, cancellation)) _cancellation = null;
            cancellation.Dispose(); IsBusy = false;
        }
    }
    [RelayCommand(CanExecute = nameof(CanCancel))]
    private void Cancel() => _cancellation?.Cancel();
    [RelayCommand(CanExecute = nameof(CanOpenOutput))]
    private async Task OpenOutputAsync(CancellationToken token) {
        if (_openOutput == null || OutputPath == null) return;
        try { await _openOutput(OutputPath, token).ConfigureAwait(true); }
        catch (Exception error) { Status = error.Message; }
    }
}
