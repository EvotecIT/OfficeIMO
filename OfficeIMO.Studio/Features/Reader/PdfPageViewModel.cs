using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// Presentation state for one virtualized PDF page. Decoded images are held only while the page is realized.
/// </summary>
public sealed partial class PdfPageViewModel : ObservableObject, IDisposable {
    private readonly PageRenderCoordinator _renderCoordinator;
    private readonly double _pageWidth;
    private readonly double _pageHeight;
    private CancellationTokenSource? _renderCancellation;
    private long _renderGeneration;
    private bool _isAttached;
    private bool _disposed;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasImage))]
    private Bitmap? _pageImage;

    [ObservableProperty]
    private bool _isRendering;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasRenderError))]
    private string? _renderError;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasDiagnostics))]
    private string? _diagnosticsSummary;

    [ObservableProperty]
    private double _displayWidth;

    [ObservableProperty]
    private double _displayHeight;

    [ObservableProperty]
    private double _renderedScale;

    private double _zoom;

    internal PdfPageViewModel(
        int pageNumber,
        double width,
        double height,
        int rotationDegrees,
        double zoom,
        PageRenderCoordinator renderCoordinator) {
        PageNumber = pageNumber;
        bool swapsAxes = Math.Abs(rotationDegrees) % 180 == 90;
        _pageWidth = Math.Max(1D, swapsAxes ? height : width);
        _pageHeight = Math.Max(1D, swapsAxes ? width : height);
        _renderCoordinator = renderCoordinator;
        _zoom = zoom;
        UpdateDisplaySize();
    }

    public int PageNumber { get; }

    public string PageLabel => $"Page {PageNumber}";

    public bool HasImage => PageImage is not null;

    public bool HasRenderError => !string.IsNullOrWhiteSpace(RenderError);

    public bool HasDiagnostics => !string.IsNullOrWhiteSpace(DiagnosticsSummary);

    internal void AttachToViewport() {
        if (_disposed || _isAttached) return;
        _isAttached = true;
        BeginRender();
    }

    internal void DetachFromViewport() {
        if (!_isAttached) return;
        _isAttached = false;
        CancelRender();
        ReplaceImage(null);
        IsRendering = false;
    }

    internal void SetZoom(double zoom) {
        if (_disposed || Math.Abs(_zoom - zoom) < 0.001D) return;
        double previousRenderScale = GetRenderScale(_zoom);
        _zoom = zoom;
        UpdateDisplaySize();
        double nextRenderScale = GetRenderScale(_zoom);

        if (Math.Abs(previousRenderScale - nextRenderScale) < 0.001D) return;
        CancelRender();
        ReplaceImage(null);
        if (_isAttached) BeginRender();
    }

    internal async Task EnsureRenderedAsync() {
        if (_disposed || !_isAttached) return;

        CancelRender();
        var cancellation = new CancellationTokenSource();
        CancellationToken token = cancellation.Token;
        _renderCancellation = cancellation;
        long generation = ++_renderGeneration;
        double requestedScale = GetRenderScale(_zoom);

        IsRendering = true;
        RenderError = null;
        DiagnosticsSummary = null;

        try {
            PdfRenderedPage rendered = await _renderCoordinator
                .GetPageAsync(PageNumber, requestedScale, token)
                .ConfigureAwait(false);

            token.ThrowIfCancellationRequested();
            using var stream = new MemoryStream(rendered.Bytes, writable: false);
            var bitmap = new Bitmap(stream);

            await Dispatcher.UIThread.InvokeAsync(() => {
                if (_disposed || !_isAttached || generation != _renderGeneration || token.IsCancellationRequested) {
                    bitmap.Dispose();
                    return;
                }

                ReplaceImage(bitmap);
                RenderedScale = rendered.Scale;
                DiagnosticsSummary = rendered.Diagnostics.Count == 0
                    ? null
                    : rendered.Diagnostics.Count == 1
                        ? "1 rendering note"
                        : $"{rendered.Diagnostics.Count} rendering notes";
            });
        } catch (OperationCanceledException) {
            // A detached page, zoom change, document close, or newer generation superseded this result.
        } catch (Exception ex) {
            await Dispatcher.UIThread.InvokeAsync(() => {
                if (!_disposed && _isAttached && generation == _renderGeneration) {
                    RenderError = ex.Message;
                }
            });
        } finally {
            await Dispatcher.UIThread.InvokeAsync(() => {
                if (!_disposed && generation == _renderGeneration) {
                    IsRendering = false;
                }
            });
        }
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _isAttached = false;
        CancelRender();
        ReplaceImage(null);
    }

    private void BeginRender() {
        _ = EnsureRenderedAsync();
    }

    private void CancelRender() {
        ++_renderGeneration;
        CancellationTokenSource? cancellation = _renderCancellation;
        _renderCancellation = null;
        if (cancellation is null) return;
        cancellation.Cancel();
        cancellation.Dispose();
    }

    private void ReplaceImage(Bitmap? replacement) {
        Bitmap? previous = PageImage;
        PageImage = replacement;
        if (!ReferenceEquals(previous, replacement)) previous?.Dispose();
    }

    private void UpdateDisplaySize() {
        DisplayWidth = Math.Round(_pageWidth * _zoom, 2);
        DisplayHeight = Math.Round(_pageHeight * _zoom, 2);
    }

    private static double GetRenderScale(double zoom) => Math.Clamp(zoom, 0.5D, 2.5D);
}
