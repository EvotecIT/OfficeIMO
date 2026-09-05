using Avalonia;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// Presentation state for one virtualized PDF page. Retained scenes and decoded fallback images are exposed only
/// while the page is realized by the viewport.
/// </summary>
public sealed partial class PdfPageViewModel : ObservableObject, IDisposable {
    private readonly PageSceneCoordinator _sceneCoordinator;
    private readonly PageRenderCoordinator _renderCoordinator;
    private readonly IStudioLocalizer _localizer;
    private readonly double _pageWidth;
    private readonly double _pageHeight;
    private CancellationTokenSource? _loadCancellation;
    private long _loadGeneration;
    private bool _isAttached;
    private bool _disposed;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasScene))]
    private PdfPageScene? _scene;

    [ObservableProperty]
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
    private string? _diagnosticsDetails;

    [ObservableProperty]
    private double _displayWidth;

    [ObservableProperty]
    private double _displayHeight;

    [ObservableProperty]
    private double _renderedScale;

    [ObservableProperty]
    private PdfEditorTool _editorTool;

    [ObservableProperty]
    private PdfEditorSelection? _selectedObject;

    [ObservableProperty]
    private PdfEditorSelectionMode _selectionMode;

    [ObservableProperty]
    private Rect? _pendingRedactionArea;

    [ObservableProperty]
    private bool _isNightMode;

    private double _zoom;

    internal PdfPageViewModel(
        int pageNumber,
        double width,
        double height,
        int rotationDegrees,
        double zoom,
        PageSceneCoordinator sceneCoordinator,
        PageRenderCoordinator renderCoordinator,
        IStudioLocalizer? localizer = null) {
        PageNumber = pageNumber;
        bool swapsAxes = Math.Abs(rotationDegrees) % 180 == 90;
        _pageWidth = Math.Max(1D, swapsAxes ? height : width);
        _pageHeight = Math.Max(1D, swapsAxes ? width : height);
        _sceneCoordinator = sceneCoordinator;
        _renderCoordinator = renderCoordinator;
        _localizer = localizer ?? new StudioLocalizer(System.Globalization.CultureInfo.GetCultureInfo("en"));
        _zoom = zoom;
        UpdateDisplaySize();
    }

    public int PageNumber { get; }

    public string PageLabel => _localizer.Format("PdfPage.Label", PageNumber);

    public bool HasScene => Scene is not null;

    public bool HasRenderError => !string.IsNullOrWhiteSpace(RenderError);

    public bool HasDiagnostics => !string.IsNullOrWhiteSpace(DiagnosticsSummary);

    internal event Action<string>? LinkActivated;

    internal event Action<PdfEditorGesture>? EditorGestureCompleted;

    internal event Action<PdfEditorSelection?>? ObjectSelected;

    internal void AttachToViewport() {
        if (_disposed || _isAttached) return;
        _isAttached = true;
        BeginLoad();
    }

    internal void DetachFromViewport() {
        if (!_isAttached) return;
        _isAttached = false;
        CancelLoad();
        Scene = null;
        ReplaceImage(null);
        IsRendering = false;
    }

    internal void SetZoom(double zoom) {
        if (_disposed || Math.Abs(_zoom - zoom) < 0.001D) return;
        double previousRenderScale = GetRenderScale(_zoom);
        _zoom = zoom;
        UpdateDisplaySize();

        if (Scene?.RequiresRasterFallback != true || Math.Abs(previousRenderScale - GetRenderScale(_zoom)) < 0.001D) {
            return;
        }

        CancelLoad();
        if (_isAttached) BeginLoad();
    }

    internal void ActivateLink(string target) {
        if (!string.IsNullOrWhiteSpace(target)) LinkActivated?.Invoke(target);
    }

    internal void CompleteEditorGesture(PdfEditorGesture gesture) => EditorGestureCompleted?.Invoke(gesture);

    internal void SelectObject(PdfEditorSelection? selection) => ObjectSelected?.Invoke(selection);

    internal async Task EnsureRenderedAsync() {
        if (_disposed || !_isAttached) return;

        CancelLoad();
        var cancellation = new CancellationTokenSource();
        CancellationToken token = cancellation.Token;
        _loadCancellation = cancellation;
        long generation = ++_loadGeneration;
        double requestedScale = GetRenderScale(_zoom);

        IsRendering = true;
        RenderError = null;
        DiagnosticsSummary = null;
        DiagnosticsDetails = null;

        try {
            PdfPageScene scene = await _sceneCoordinator
                .GetPageAsync(PageNumber, token)
                .ConfigureAwait(false);
            token.ThrowIfCancellationRequested();

            await Dispatcher.UIThread.InvokeAsync(() => {
                if (!_disposed && _isAttached && generation == _loadGeneration && !token.IsCancellationRequested) {
                    Scene = scene;
                }
            });
            token.ThrowIfCancellationRequested();

            Bitmap? bitmap = null;
            IReadOnlyList<string> diagnostics = scene.Diagnostics;
            if (scene.RequiresRasterFallback) {
                PdfRenderedPage rendered = await _renderCoordinator
                    .GetPageAsync(PageNumber, requestedScale, token)
                    .ConfigureAwait(false);
                token.ThrowIfCancellationRequested();
                using var stream = new MemoryStream(rendered.Bytes, writable: false);
                bitmap = new Bitmap(stream);
                diagnostics = diagnostics.Concat(rendered.Diagnostics).Distinct(StringComparer.Ordinal).ToArray();
            }

            await Dispatcher.UIThread.InvokeAsync(() => {
                if (_disposed || !_isAttached || generation != _loadGeneration || token.IsCancellationRequested) {
                    bitmap?.Dispose();
                    return;
                }

                Scene = scene;
                ReplaceImage(bitmap);
                RenderedScale = scene.RequiresRasterFallback ? requestedScale : 0D;
                DiagnosticsSummary = diagnostics.Count == 0 ? null : diagnostics[0];
                DiagnosticsDetails = diagnostics.Count == 0
                    ? null
                    : string.Join(Environment.NewLine, diagnostics);
            });
        } catch (OperationCanceledException) {
            // A detached page, zoom change, document close, or newer generation superseded this result.
        } catch (Exception ex) {
            await Dispatcher.UIThread.InvokeAsync(() => {
                if (!_disposed && _isAttached && generation == _loadGeneration) RenderError = ex.Message;
            });
        } finally {
            await Dispatcher.UIThread.InvokeAsync(() => {
                if (!_disposed && generation == _loadGeneration) IsRendering = false;
            });
        }
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _isAttached = false;
        CancelLoad();
        Scene = null;
        ReplaceImage(null);
    }

    private void BeginLoad() => _ = EnsureRenderedAsync();

    private void CancelLoad() {
        ++_loadGeneration;
        CancellationTokenSource? cancellation = _loadCancellation;
        _loadCancellation = null;
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
