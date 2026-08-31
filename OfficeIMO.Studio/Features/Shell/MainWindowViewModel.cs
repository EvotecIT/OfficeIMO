using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickPdf;
    private PdfDocumentSession? _session;
    private PageRenderCoordinator? _renderCoordinator;
    private CancellationTokenSource? _openCancellation;
    private double _viewportWidth = 1000D;
    private double _viewportHeight = 700D;
    private ViewerZoomMode _zoomMode = ViewerZoomMode.FitWidth;
    private bool _disposed;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsEmpty))]
    private bool _hasDocument;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsEmpty))]
    private bool _isOpening;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasError))]
    private string? _errorMessage;

    [ObservableProperty]
    private string _documentName = "OfficeIMO Studio";

    [ObservableProperty]
    private string _documentDescription = "Open a PDF to begin";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(SelectedPagePosition))]
    [NotifyPropertyChangedFor(nameof(CanGoPrevious))]
    [NotifyPropertyChangedFor(nameof(CanGoNext))]
    private PdfPageViewModel? _selectedPage;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ZoomLabel))]
    private double _zoom = 1D;

    internal MainWindowViewModel(Func<CancellationToken, Task<string?>> pickPdf) {
        _pickPdf = pickPdf ?? throw new ArgumentNullException(nameof(pickPdf));
    }

    public ObservableCollection<PdfPageViewModel> Pages { get; } = new();

    public bool IsEmpty => !HasDocument && !IsOpening;

    public bool HasError => !string.IsNullOrWhiteSpace(ErrorMessage);

    public string SelectedPagePosition => SelectedPage is null
        ? "No page"
        : $"Page {SelectedPage.PageNumber} of {Pages.Count}";

    public string ZoomLabel => $"{Zoom:P0}";

    public bool CanGoPrevious => SelectedPage is { PageNumber: > 1 };

    public bool CanGoNext => SelectedPage is not null && SelectedPage.PageNumber < Pages.Count;

    partial void OnSelectedPageChanged(PdfPageViewModel? value) {
        if (value is not null && _zoomMode != ViewerZoomMode.Custom) {
            ApplyFitZoom();
        }
    }

    [RelayCommand]
    private async Task OpenAsync(CancellationToken cancellationToken) {
        string? path = await _pickPdf(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) {
            await OpenDocumentAsync(path, cancellationToken).ConfigureAwait(true);
        }
    }

    internal async Task OpenDocumentAsync(string path, CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        _openCancellation?.Cancel();
        _openCancellation?.Dispose();
        _openCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        CancellationTokenSource currentCancellation = _openCancellation;

        IsOpening = true;
        ErrorMessage = null;

        try {
            PdfDocumentSession session = await PdfDocumentSession
                .OpenAsync(path, currentCancellation.Token)
                .ConfigureAwait(true);
            currentCancellation.Token.ThrowIfCancellationRequested();

            var coordinator = new PageRenderCoordinator(session.RenderPageAsync);
            var pages = session.Pages
                .Select(page => new PdfPageViewModel(
                    page.PageNumber,
                    page.Width,
                    page.Height,
                    page.RotationDegrees,
                    Zoom,
                    coordinator))
                .ToArray();

            if (!ReferenceEquals(currentCancellation, _openCancellation)) {
                foreach (PdfPageViewModel page in pages) page.Dispose();
                coordinator.Dispose();
                return;
            }

            ReplaceDocument(session, coordinator, pages);
        } catch (OperationCanceledException) when (currentCancellation.IsCancellationRequested) {
            // A newer open request, document close, or application shutdown superseded this operation.
        } catch (Exception ex) {
            if (ReferenceEquals(currentCancellation, _openCancellation)) {
                ErrorMessage = ex.Message;
            }
        } finally {
            if (ReferenceEquals(currentCancellation, _openCancellation)) {
                IsOpening = false;
                _openCancellation = null;
                currentCancellation.Dispose();
            }
        }
    }

    [RelayCommand]
    private void CloseDocument() {
        _openCancellation?.Cancel();
        ReplaceDocument(null, null, Array.Empty<PdfPageViewModel>());
        ErrorMessage = null;
    }

    [RelayCommand]
    private void DismissError() {
        ErrorMessage = null;
    }

    [RelayCommand]
    private void PreviousPage() {
        if (!CanGoPrevious || SelectedPage is null) return;
        SelectedPage = Pages[SelectedPage.PageNumber - 2];
    }

    [RelayCommand]
    private void NextPage() {
        if (!CanGoNext || SelectedPage is null) return;
        SelectedPage = Pages[SelectedPage.PageNumber];
    }

    [RelayCommand]
    private void ZoomIn() {
        _zoomMode = ViewerZoomMode.Custom;
        ApplyZoom(Math.Min(3D, Zoom + 0.25D));
    }

    [RelayCommand]
    private void ZoomOut() {
        _zoomMode = ViewerZoomMode.Custom;
        ApplyZoom(Math.Max(0.25D, Zoom - 0.25D));
    }

    [RelayCommand]
    private void ActualSize() {
        _zoomMode = ViewerZoomMode.Custom;
        ApplyZoom(1D);
    }

    [RelayCommand]
    private void FitWidth() {
        _zoomMode = ViewerZoomMode.FitWidth;
        ApplyFitZoom();
    }

    [RelayCommand]
    private void FitPage() {
        _zoomMode = ViewerZoomMode.FitPage;
        ApplyFitZoom();
    }

    internal void SetViewportSize(double width, double height) {
        if (width <= 0 || height <= 0) return;
        _viewportWidth = width;
        _viewportHeight = height;
        if (_zoomMode != ViewerZoomMode.Custom) ApplyFitZoom();
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _openCancellation?.Cancel();
        _openCancellation?.Dispose();
        _openCancellation = null;
        ReplaceDocument(null, null, Array.Empty<PdfPageViewModel>());
    }

    private void ReplaceDocument(
        PdfDocumentSession? session,
        PageRenderCoordinator? coordinator,
        IReadOnlyList<PdfPageViewModel> pages) {
        foreach (PdfPageViewModel page in Pages) page.Dispose();
        Pages.Clear();
        _renderCoordinator?.Dispose();

        _session = session;
        _renderCoordinator = coordinator;

        foreach (PdfPageViewModel page in pages) Pages.Add(page);

        HasDocument = session is not null;
        DocumentName = session?.FileName ?? "OfficeIMO Studio";
        DocumentDescription = session is null
            ? "Open a PDF to begin"
            : $"{session.Pages.Count:N0} {(session.Pages.Count == 1 ? "page" : "pages")} · {FormatByteSize(session.FileSize)}";
        SelectedPage = Pages.FirstOrDefault();
        OnPropertyChanged(nameof(SelectedPagePosition));

        if (session is not null) ApplyFitZoom();
    }

    private void ApplyFitZoom() {
        if (Pages.Count == 0) return;
        PdfPageViewModel page = SelectedPage ?? Pages[0];
        double unscaledWidth = page.DisplayWidth / Math.Max(Zoom, 0.01D);
        double unscaledHeight = page.DisplayHeight / Math.Max(Zoom, 0.01D);
        double availableWidth = Math.Max(200D, _viewportWidth - 72D);
        double availableHeight = Math.Max(200D, _viewportHeight - 72D);
        double target = _zoomMode == ViewerZoomMode.FitPage
            ? Math.Min(availableWidth / unscaledWidth, availableHeight / unscaledHeight)
            : availableWidth / unscaledWidth;
        ApplyZoom(Math.Clamp(target, 0.25D, 3D));
    }

    private void ApplyZoom(double zoom) {
        zoom = Math.Round(zoom, 2);
        if (Math.Abs(Zoom - zoom) < 0.001D) return;
        Zoom = zoom;
        foreach (PdfPageViewModel page in Pages) page.SetZoom(zoom);
    }

    private static string FormatByteSize(long bytes) {
        string[] units = ["B", "KB", "MB", "GB"];
        double value = bytes;
        int unit = 0;
        while (value >= 1024D && unit < units.Length - 1) {
            value /= 1024D;
            unit++;
        }

        return $"{value:0.#} {units[unit]}";
    }

    private void ThrowIfDisposed() {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }

    private enum ViewerZoomMode {
        Custom,
        FitWidth,
        FitPage
    }
}
