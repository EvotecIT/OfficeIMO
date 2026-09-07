using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Drawing;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Editable recognized text alongside a bounded, selected-frame source preview.</summary>
public sealed partial class ImageOcrReviewViewModel : ObservableObject, IDisposable {
    private readonly byte[] _image;
    private readonly Action _cancel;
    private readonly TaskCompletionSource<string> _completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private CancellationTokenSource? _previewCancellation;
    private bool _disposed;
    internal Task<string> Completion => _completion.Task;
    internal Task PreviewTask { get; private set; } = Task.CompletedTask;
    internal ImageOcrReviewViewModel(ImageOcrWorkflowReview review, IStudioLocalizer localizer, Action cancel) {
        _image = review.GetImageBytes(); _cancel = cancel; _text = review.Text;
        SourceName = review.SourceName;
        Details = string.Join(" · ", review.Recognition.Recognitions.Select(item =>
            $"{item.Result.Provider} · {item.Result.Language} · {item.Result.Confidence:P0}"));
        Diagnostics = string.Join(Environment.NewLine, review.Recognition.Diagnostics.Select(item => item.Message));
        if (!OfficeRasterContainerInspector.TryInspect(_image, out var container) || container is null)
            throw new InvalidDataException(localizer.GetOrDefault("ImageOcr.PreviewUnsupported", "The source image cannot be previewed safely."));
        Frames = Enumerable.Range(1, container.Count).Select(number => new OcrReviewPageChoice(number,
            localizer.FormatOrDefault("ImageOcr.Frame", "Image {0}", number))).ToArray();
        SelectedFrame = Frames.First();
    }
    public string SourceName { get; }
    public string Details { get; }
    public string Diagnostics { get; }
    public IReadOnlyList<OcrReviewPageChoice> Frames { get; }
    [ObservableProperty] private string _text;
    [ObservableProperty] private Bitmap? _preview;
    [ObservableProperty] private bool _isZoomed;
    [ObservableProperty] private OcrReviewPageChoice? _selectedFrame;
    [ObservableProperty] private string? _previewError;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CommitCommand))]
    private bool _isLoading;
    private bool CanCommit => !_disposed && !IsLoading && Preview is not null && PreviewError is null && !_completion.Task.IsCompleted;
    partial void OnSelectedFrameChanged(OcrReviewPageChoice? value) {
        if (!_disposed && value is not null) PreviewTask = LoadPreviewAsync(value.Number - 1);
    }
    private async Task LoadPreviewAsync(int frame) {
        _previewCancellation?.Cancel();
        using var operation = new CancellationTokenSource();
        _previewCancellation = operation;
        Preview?.Dispose(); Preview = null; PreviewError = null; IsLoading = true;
        try {
            byte[] png = await Task.Run(() => {
                if (!OfficeRasterImageDecoder.TryDecode(_image, new OfficeRasterDecodeOptions {
                    FrameIndex = frame, MaximumDecodedPixels = 20_000_000, CancellationToken = operation.Token
                }, out var raster, out var info) || raster is null) throw new InvalidDataException(info.Diagnostic);
                return OfficeRasterImageEncoder.Encode(raster, OfficeImageExportFormat.Png, null, 100L * 1024 * 1024, operation.Token);
            }, operation.Token).ConfigureAwait(true);
            operation.Token.ThrowIfCancellationRequested();
            if (_disposed || !ReferenceEquals(_previewCancellation, operation)) return;
            using var stream = new MemoryStream(png, writable: false);
            Preview = new Bitmap(stream);
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) { }
        catch (Exception error) {
            if (!_disposed && ReferenceEquals(_previewCancellation, operation)) PreviewError = error.Message;
        } finally {
            if (ReferenceEquals(_previewCancellation, operation)) {
                _previewCancellation = null; IsLoading = false; CommitCommand.NotifyCanExecuteChanged();
            }
        }
    }
    [RelayCommand(CanExecute = nameof(CanCommit))]
    private void Commit() {
        if (CanCommit) _completion.TrySetResult(Text);
        CommitCommand.NotifyCanExecuteChanged();
    }
    [RelayCommand] private void Cancel() { if (!_disposed) _cancel(); }
    public void Dispose() {
        if (_disposed) return;
        _disposed = true; _previewCancellation?.Cancel(); Preview?.Dispose(); Preview = null;
        _completion.TrySetCanceled(); CommitCommand.NotifyCanExecuteChanged();
    }
}
