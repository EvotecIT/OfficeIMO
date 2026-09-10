using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionWorkbenchViewModel {
    private CancellationTokenSource? _previewCancellation;
    public ObservableCollection<Bitmap> OutputPreviewPages { get; } = new();
    [ObservableProperty] private string _outputPreviewStatus = string.Empty;

    partial void OnSelectedJobChanged(ConversionJobViewModel? value) => ClearOutputPreview();

    private void ClearOutputPreview() {
        _previewCancellation?.Cancel();
        foreach (Bitmap page in OutputPreviewPages) page.Dispose();
        OutputPreviewPages.Clear();
        OutputPreviewStatus = string.Empty;
    }

    [RelayCommand]
    private async Task PreviewOutputAsync(CancellationToken token) {
        if (SelectedJob is not { State: ConversionJobState.Completed, OutputPath: { } path } job) return;
        ClearOutputPreview();
        using var cancellation = CancellationTokenSource.CreateLinkedTokenSource(token);
        _previewCancellation = cancellation;
        OutputPreviewStatus = T("Preview.Loading", "Preparing a sample of the saved output…");
        try {
            byte[] bytes;
            if (_storage is not null) bytes = (await _storage.ReadSnapshotAsync(path, cancellation.Token, 128L * 1024 * 1024).ConfigureAwait(true)).Bytes;
            else {
                using var storage = new Infrastructure.StudioStorageAccess();
                bytes = (await storage.ReadSnapshotAsync(path, cancellation.Token, 128L * 1024 * 1024).ConfigureAwait(true)).Bytes;
            }
            OfficeWorkflowDocumentPreview preview = await Task.Run(() => OfficeWorkflowRunner.PreviewDocument(bytes,
                job.Route.Route.TargetExtension, cancellation.Token), cancellation.Token).ConfigureAwait(true);
            cancellation.Token.ThrowIfCancellationRequested();
            if (!ReferenceEquals(SelectedJob, job) || job.OutputPath != path) return;
            foreach (var page in preview.Pages) {
                using var stream = new MemoryStream(page.Bytes!);
                OutputPreviewPages.Add(new Bitmap(stream));
            }
            string[] warnings = preview.Diagnostics.Where(item => item.Severity == OfficeWorkflowDiagnosticSeverity.Warning).Select(item => item.Message).Distinct().ToArray();
            OutputPreviewStatus = T("Preview.Sample", "First pages of the saved artifact, rendered by OfficeIMO. Open the output to review the whole document.") +
                (warnings.Length == 0 ? string.Empty : " " + string.Join(" ", warnings));
        } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { }
        catch (Exception error) {
            if (ReferenceEquals(SelectedJob, job)) OutputPreviewStatus = T("Preview.Failed", "Preview unavailable: ") + error.Message;
        } finally {
            if (ReferenceEquals(_previewCancellation, cancellation)) _previewCancellation = null;
        }
    }

    [RelayCommand]
    private async Task OpenOutputAsync(CancellationToken token) {
        if (_openOutput is null || SelectedJob is not { State: ConversionJobState.Completed, OutputPath: { } path }) return;
        try { await _openOutput(path, token).ConfigureAwait(true); }
        catch (Exception error) { Status = T("Output.OpenFailed", "The output could not be opened: ") + error.Message; }
    }
}
