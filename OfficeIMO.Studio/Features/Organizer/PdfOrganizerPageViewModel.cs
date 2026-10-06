using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PdfOrganizerPageViewModel : ObservableObject, IDisposable {
    private readonly PageSceneCoordinator _sceneCoordinator;
    private readonly PageRenderCoordinator _renderCoordinator;
    private readonly IStudioLocalizer _localizer;
    private CancellationTokenSource? _loadCancellation;
    private long _loadGeneration;
    private double _renderScaling = 1D;
    private bool _attached;
    private bool _disposed;

    [ObservableProperty]
    private PdfPageScene? _scene;

    [ObservableProperty]
    private Bitmap? _fallbackImage;

    [ObservableProperty]
    private bool _isLoading;

    [ObservableProperty]
    private bool _isSelected;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasError))]
    private string? _error;

    internal PdfOrganizerPageViewModel(
        int pageNumber,
        double width,
        double height,
        int rotationDegrees,
        PageSceneCoordinator sceneCoordinator,
        PageRenderCoordinator renderCoordinator,
        IStudioLocalizer? localizer = null) {
        PageNumber = pageNumber;
        RotationDegrees = rotationDegrees;
        bool swapsAxes = Math.Abs(rotationDegrees) % 180 == 90;
        double visualWidth = Math.Max(1D, swapsAxes ? height : width);
        double visualHeight = Math.Max(1D, swapsAxes ? width : height);
        double scale = Math.Min(112D / visualWidth, 142D / visualHeight);
        ThumbnailWidth = Math.Round(visualWidth * scale, 2);
        ThumbnailHeight = Math.Round(visualHeight * scale, 2);
        _sceneCoordinator = sceneCoordinator;
        _renderCoordinator = renderCoordinator;
        _localizer = localizer ?? new StudioLocalizer(System.Globalization.CultureInfo.GetCultureInfo("en"));
    }

    public int PageNumber { get; }

    public int RotationDegrees { get; }

    public string Label => _localizer.Format("PdfPage.Label", PageNumber);

    public double ThumbnailWidth { get; }

    public double ThumbnailHeight { get; }

    public bool HasError => !string.IsNullOrWhiteSpace(Error);

    /// <summary>Applies the presenting display's scaling and reloads a raster thumbnail whose resolution changes.</summary>
    internal void SetRenderScaling(double renderScaling) {
        renderScaling = PdfRasterScale.NormalizeRenderScaling(renderScaling);
        if (_disposed || Math.Abs(_renderScaling - renderScaling) < 0.001D) return;
        double previousScale = GetRenderScale(Scene, _renderScaling);
        _renderScaling = renderScaling;
        if (!_attached || Scene?.RequiresRasterFallback != true || Math.Abs(previousScale - GetRenderScale(Scene, renderScaling)) < 0.001D) return;
        _ = LoadAsync();
    }

    /// <summary>The raster scale that fills the thumbnail box at the current display scaling.</summary>
    internal double GetRenderScale(PdfPageScene? scene) => GetRenderScale(scene, _renderScaling);

    private double GetRenderScale(PdfPageScene? scene, double renderScaling) {
        double drawingWidth = scene?.Drawing.Width ?? ThumbnailWidth;
        double drawingHeight = scene?.Drawing.Height ?? ThumbnailHeight;
        double fit = Math.Min(ThumbnailWidth / Math.Max(1D, drawingWidth), ThumbnailHeight / Math.Max(1D, drawingHeight));
        return PdfRasterScale.Compose(
            fit,
            renderScaling,
            drawingWidth,
            drawingHeight,
            StudioPdfSecurityPolicy.MaximumRasterPixels);
    }

    internal void Attach() {
        if (_disposed || _attached) return;
        _attached = true;
        _ = LoadAsync();
    }

    internal void Detach() {
        if (!_attached) return;
        _attached = false;
        CancelLoad();
        Scene = null;
        ReplaceImage(null);
        IsLoading = false;
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _attached = false;
        CancelLoad();
        Scene = null;
        ReplaceImage(null);
    }

    private async Task LoadAsync() {
        Dispatcher uiDispatcher = Dispatcher.UIThread;
        CancelLoad();
        long generation = ++_loadGeneration;
        var cancellation = new CancellationTokenSource();
        _loadCancellation = cancellation;
        CancellationToken token = cancellation.Token;
        double renderScaling = _renderScaling;
        IsLoading = true;
        Error = null;
        try {
            PdfPageScene scene = await _sceneCoordinator.GetPageAsync(PageNumber, token).ConfigureAwait(false);
            Bitmap? image = null;
            if (scene.RequiresRasterFallback) {
                PdfRenderedPage rendered = await _renderCoordinator.GetPageAsync(PageNumber, GetRenderScale(scene, renderScaling), token).ConfigureAwait(false);
                using var stream = new MemoryStream(rendered.Bytes, writable: false);
                image = new Bitmap(stream);
            }

            await uiDispatcher.InvokeAsync(() => {
                if (!_disposed && _attached && generation == _loadGeneration && !token.IsCancellationRequested) {
                    Scene = scene;
                    ReplaceImage(image);
                } else {
                    image?.Dispose();
                }
            });
        } catch (OperationCanceledException) {
            // Virtualization or a document refresh superseded this load.
        } catch (Exception ex) {
            await uiDispatcher.InvokeAsync(() => {
                if (!_disposed && _attached && generation == _loadGeneration) Error = ex.Message;
            });
        } finally {
            if (!_disposed) {
                await uiDispatcher.InvokeAsync(() => {
                    if (!_disposed && generation == _loadGeneration) IsLoading = false;
                });
            }
        }
    }

    private void CancelLoad() {
        _loadGeneration++;
        CancellationTokenSource? cancellation = _loadCancellation;
        _loadCancellation = null;
        if (cancellation is null) return;
        cancellation.Cancel();
        cancellation.Dispose();
    }

    private void ReplaceImage(Bitmap? replacement) {
        Bitmap? previous = FallbackImage;
        FallbackImage = replacement;
        if (!ReferenceEquals(previous, replacement)) previous?.Dispose();
    }
}
