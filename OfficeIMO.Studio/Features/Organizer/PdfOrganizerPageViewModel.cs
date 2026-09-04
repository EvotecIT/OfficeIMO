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
        CancelLoad();
        long generation = ++_loadGeneration;
        var cancellation = new CancellationTokenSource();
        _loadCancellation = cancellation;
        CancellationToken token = cancellation.Token;
        IsLoading = true;
        Error = null;
        try {
            PdfPageScene scene = await _sceneCoordinator.GetPageAsync(PageNumber, token).ConfigureAwait(false);
            Bitmap? image = null;
            if (scene.RequiresRasterFallback) {
                PdfRenderedPage rendered = await _renderCoordinator.GetPageAsync(PageNumber, 0.25D, token).ConfigureAwait(false);
                using var stream = new MemoryStream(rendered.Bytes, writable: false);
                image = new Bitmap(stream);
            }

            await Dispatcher.UIThread.InvokeAsync(() => {
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
            await Dispatcher.UIThread.InvokeAsync(() => {
                if (!_disposed && _attached && generation == _loadGeneration) Error = ex.Message;
            });
        } finally {
            await Dispatcher.UIThread.InvokeAsync(() => {
                if (!_disposed && generation == _loadGeneration) IsLoading = false;
            });
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
