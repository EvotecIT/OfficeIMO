using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Features.Workflows;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickPdf;
    private readonly Func<CancellationToken, Task<string?>> _pickSavePdf;
    private readonly Func<CancellationToken, Task<IReadOnlyList<string>>> _pickImportPdfs;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputFolder;
    private readonly Func<CancellationToken, Task<string?>> _pickImage;
    private readonly Func<Uri, Task> _openUri;
    private readonly Func<Task<UnsavedChangesDecision>> _confirmUnsavedChanges;
    private readonly Func<int, Task<bool>> _confirmPageDeletion;
    private PdfWorkspace? _workspace;
    private PdfDocumentSession? _session;
    private PageSceneCoordinator? _sceneCoordinator;
    private PageRenderCoordinator? _renderCoordinator;
    private CancellationTokenSource? _openCancellation;
    private double _viewportWidth = 1000D;
    private double _viewportHeight = 700D;
    private ViewerZoomMode _zoomMode = ViewerZoomMode.FitWidth;
    private bool _disposed;
    private bool _discardOnNextTransition;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsEmpty))]
    [NotifyPropertyChangedFor(nameof(ShowPdfDocumentControls))]
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

    internal MainWindowViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>>? pickSavePdf = null,
        Func<CancellationToken, Task<IReadOnlyList<string>>>? pickImportPdfs = null,
        Func<CancellationToken, Task<string?>>? pickOutputFolder = null,
        Func<Uri, Task>? openUri = null,
        Func<Task<UnsavedChangesDecision>>? confirmUnsavedChanges = null,
        Func<CancellationToken, Task<string?>>? pickImage = null,
        Func<int, Task<bool>>? confirmPageDeletion = null,
        Func<CancellationToken, Task<IReadOnlyList<string>>>? pickWorkflowFiles = null) {
        _pickPdf = pickPdf ?? throw new ArgumentNullException(nameof(pickPdf));
        _pickSavePdf = pickSavePdf ?? (_ => Task.FromResult<string?>(null));
        _pickImportPdfs = pickImportPdfs ?? (_ => Task.FromResult<IReadOnlyList<string>>(Array.Empty<string>()));
        _pickOutputFolder = pickOutputFolder ?? (_ => Task.FromResult<string?>(null));
        _pickImage = pickImage ?? (_ => Task.FromResult<string?>(null));
        _openUri = openUri ?? (_ => Task.CompletedTask);
        _confirmUnsavedChanges = confirmUnsavedChanges ?? (() => Task.FromResult(UnsavedChangesDecision.Discard));
        _confirmPageDeletion = confirmPageDeletion ?? (_ => Task.FromResult(false));
        ConversionWorkbench = new ConversionWorkbenchViewModel(
            pickWorkflowFiles ?? (_ => Task.FromResult<IReadOnlyList<string>>(Array.Empty<string>())),
            _pickOutputFolder);
        DocumentHealth = new DocumentHealthViewModel(_pickPdf, _pickOutputFolder);
        ConversionWorkbench.PropertyChanged += OnWorkflowPropertyChanged;
        DocumentHealth.PropertyChanged += OnWorkflowPropertyChanged;
    }

    public ObservableCollection<PdfPageViewModel> Pages { get; } = new();

    public ObservableCollection<PdfOrganizerPageViewModel> OrganizerPages { get; } = new();

    public bool IsEmpty => !HasDocument && !IsOpening;

    public bool HasError => !string.IsNullOrWhiteSpace(ErrorMessage);

    public string SelectedPagePosition => SelectedPage is null
        ? "No page"
        : $"Page {SelectedPage.PageNumber} of {Pages.Count}";

    public string ZoomLabel => $"{Zoom:P0}";

    public bool CanGoPrevious => SelectedPage is { PageNumber: > 1 };

    public bool CanGoNext => SelectedPage is not null && SelectedPage.PageNumber < Pages.Count;

    public bool IsDirty => _workspace?.IsDirty == true;

    public bool CanUndo => _workspace?.CanUndo == true;

    public bool CanRedo => _workspace?.CanRedo == true;

    public bool HasRecovery => _workspace?.HasRecovery == true;

    public bool CanMutatePages => _workspace?.CanMutatePages == true;

    public bool CanExtractPages => _workspace?.CanExtractPages == true;

    public bool CanImportPages => _workspace?.CanImportPages == true;

    public bool HasSecurityWarning => !string.IsNullOrWhiteSpace(SecurityWarning);

    public string? SecurityWarning => _workspace?.SecurityWarning;

    public bool CanStartDocumentTransition => !IsWorkspaceBusy && !IsOpening;

    public bool CanCancelOperation => IsWorkspaceBusy || IsOpening || ConversionWorkbench.IsBusy || DocumentHealth.IsBusy;

    partial void OnIsOpeningChanged(bool value) {
        OnPropertyChanged(nameof(CanStartDocumentTransition));
        OnPropertyChanged(nameof(CanCancelOperation));
    }

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
        if (!await PrepareDocumentTransitionAsync().ConfigureAwait(true)) return;
        _openCancellation?.Cancel();
        _openCancellation?.Dispose();
        _openCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        CancellationTokenSource currentCancellation = _openCancellation;

        IsOpening = true;
        ErrorMessage = null;
        PdfWorkspace? candidateWorkspace = null;
        PageSceneCoordinator? candidateSceneCoordinator = null;
        PageRenderCoordinator? candidateRenderCoordinator = null;
        PdfPageViewModel[] candidatePages = [];
        PdfOrganizerPageViewModel[] candidateOrganizerPages = [];
        bool installed = false;

        try {
            candidateWorkspace = await PdfWorkspace
                .OpenAsync(path, currentCancellation.Token)
                .ConfigureAwait(true);
            currentCancellation.Token.ThrowIfCancellationRequested();

            PdfDocumentSession session = PdfDocumentSession.FromWorkspace(candidateWorkspace);

            candidateSceneCoordinator = new PageSceneCoordinator(session.LoadPageSceneAsync);
            candidateRenderCoordinator = new PageRenderCoordinator(session.RenderPageAsync);
            candidatePages = session.Pages
                .Select(page => new PdfPageViewModel(
                    page.PageNumber,
                    page.Width,
                    page.Height,
                    page.RotationDegrees,
                    Zoom,
                    candidateSceneCoordinator,
                    candidateRenderCoordinator))
                .ToArray();
            candidateOrganizerPages = session.Pages
                .Select(page => new PdfOrganizerPageViewModel(
                    page.PageNumber,
                    page.Width,
                    page.Height,
                    page.RotationDegrees,
                    candidateSceneCoordinator,
                    candidateRenderCoordinator))
                .ToArray();

            if (!ReferenceEquals(currentCancellation, _openCancellation)) return;

            ReplaceDocument(
                candidateWorkspace,
                session,
                candidateSceneCoordinator,
                candidateRenderCoordinator,
                candidatePages,
                candidateOrganizerPages);
            WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
            installed = true;
        } catch (OperationCanceledException) when (currentCancellation.IsCancellationRequested) {
            // A newer open request, document close, or application shutdown superseded this operation.
            _discardOnNextTransition = false;
        } catch (Exception ex) {
            _discardOnNextTransition = false;
            if (ReferenceEquals(currentCancellation, _openCancellation)) {
                ErrorMessage = ex.Message;
            }
        } finally {
            if (!installed) {
                foreach (PdfPageViewModel page in candidatePages) page.Dispose();
                foreach (PdfOrganizerPageViewModel page in candidateOrganizerPages) page.Dispose();
                candidateSceneCoordinator?.Dispose();
                candidateRenderCoordinator?.Dispose();
                candidateWorkspace?.Dispose();
            }
            if (ReferenceEquals(currentCancellation, _openCancellation)) {
                IsOpening = false;
                _openCancellation = null;
                currentCancellation.Dispose();
            }
        }
    }

    [RelayCommand]
    private async Task CloseDocumentAsync() {
        await RequestCloseDocumentAsync().ConfigureAwait(true);
    }

    internal async Task<bool> RequestCloseDocumentAsync() {
        if (!await PrepareDocumentTransitionAsync().ConfigureAwait(true)) return false;
        _openCancellation?.Cancel();
        ReplaceDocument(null, null, null, null, Array.Empty<PdfPageViewModel>(), Array.Empty<PdfOrganizerPageViewModel>());
        ErrorMessage = null;
        return true;
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
        ConversionWorkbench.PropertyChanged -= OnWorkflowPropertyChanged;
        DocumentHealth.PropertyChanged -= OnWorkflowPropertyChanged;
        ConversionWorkbench.Dispose();
        DocumentHealth.Dispose();
        CancelCurrentOperation();
        if (IsWorkspaceBusy) {
            _disposeWhenIdle = true;
            return;
        }
        _openCancellation?.Cancel();
        _openCancellation?.Dispose();
        _openCancellation = null;
        ReplaceDocument(null, null, null, null, Array.Empty<PdfPageViewModel>(), Array.Empty<PdfOrganizerPageViewModel>());
    }

    private async Task<bool> PrepareDocumentTransitionAsync() {
        if (IsWorkspaceBusy || IsOpening) {
            OperationStatus = "Cancel or wait for the current operation before changing documents.";
            return false;
        }
        if (!IsDirty) return true;

        UnsavedChangesDecision decision = await _confirmUnsavedChanges().ConfigureAwait(true);
        if (decision == UnsavedChangesDecision.Save) {
            await RunSaveAsync(path: null, CancellationToken.None).ConfigureAwait(true);
            return !IsDirty;
        }
        if (decision == UnsavedChangesDecision.Discard) {
            _discardOnNextTransition = true;
            return true;
        }
        return false;
    }

    private void ReplaceDocument(
        PdfWorkspace? workspace,
        PdfDocumentSession? session,
        PageSceneCoordinator? sceneCoordinator,
        PageRenderCoordinator? renderCoordinator,
        IReadOnlyList<PdfPageViewModel> pages,
        IReadOnlyList<PdfOrganizerPageViewModel> organizerPages,
        IReadOnlyCollection<int>? organizerSelection = null) {
        bool isDocumentTransition = !ReferenceEquals(_workspace, workspace);
        CancelPendingRedaction();
        ClearAnnotationSelection();
        foreach (PdfPageViewModel page in Pages) page.Dispose();
        foreach (PdfOrganizerPageViewModel page in OrganizerPages) page.Dispose();
        Pages.Clear();
        OrganizerPages.Clear();
        _organizerSelection.Clear();
        _sceneCoordinator?.Dispose();
        _renderCoordinator?.Dispose();
        if (!ReferenceEquals(_workspace, workspace)) {
            if (_discardOnNextTransition) _workspace?.DiscardRecovery();
            _workspace?.Dispose();
            _discardOnNextTransition = false;
        }

        _workspace = workspace;
        _session = session;
        _sceneCoordinator = sceneCoordinator;
        _renderCoordinator = renderCoordinator;

        if (isDocumentTransition) {
            SearchQuery = string.Empty;
            SearchResults.Clear();
            SelectedSearchResult = null;
            SelectedBookmark = null;
            OperationStatus = null;
            OperationProgressFraction = 0D;
        }

        foreach (PdfPageViewModel page in pages) {
            page.LinkActivated += OnPageLinkActivated;
            page.EditorGestureCompleted += OnPageEditorGestureCompleted;
            page.AnnotationSelected += OnPageAnnotationSelected;
            page.EditorTool = ActiveEditorTool;
            Pages.Add(page);
        }
        if (organizerSelection is not null) {
            foreach (int pageNumber in organizerSelection.Where(pageNumber => pageNumber >= 1 && pageNumber <= organizerPages.Count)) {
                _organizerSelection.Add(pageNumber);
            }
        }
        foreach (PdfOrganizerPageViewModel page in organizerPages) {
            page.IsSelected = _organizerSelection.Contains(page.PageNumber);
            OrganizerPages.Add(page);
        }

        HasDocument = session is not null;
        DocumentName = session?.FileName ?? "OfficeIMO Studio";
        DocumentDescription = session is null
            ? "Open a PDF to begin"
            : $"{session.Pages.Count:N0} {(session.Pages.Count == 1 ? "page" : "pages")} · {FormatByteSize(session.FileSize)}";
        SelectedPage = Pages.FirstOrDefault();
        OnPropertyChanged(nameof(SelectedPagePosition));
        OnPropertyChanged(nameof(HasOrganizerSelection));
        OnPropertyChanged(nameof(CanDeleteSelection));
        OnPropertyChanged(nameof(OrganizerSelectionLabel));
        NotifyWorkspaceStateChanged();

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
