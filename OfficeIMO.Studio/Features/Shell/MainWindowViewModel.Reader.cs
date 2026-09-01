using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private PdfDocumentSession? _comparisonSession;
    private PageSceneCoordinator? _comparisonSceneCoordinator;
    private PageRenderCoordinator? _comparisonRenderCoordinator;
    private CancellationTokenSource? _comparisonCancellation;
    private ReaderLayoutChoice? _layoutBeforeComparison;
    private bool _synchronizingComparison;

    public IReadOnlyList<ReaderLayoutChoice> ReaderLayoutChoices { get; } = [
        new(ReaderLayoutMode.SinglePage, "Single page", "Show only the selected page."),
        new(ReaderLayoutMode.Continuous, "Continuous", "Scroll through pages vertically."),
        new(ReaderLayoutMode.TwoPage, "Two page", "Read a cover page followed by paired spreads."),
        new(ReaderLayoutMode.Grid, "Grid", "Browse the whole document as a page grid.")
    ];

    public ObservableCollection<PdfPageViewModel> ReaderPages { get; } = new();

    public ObservableCollection<ReaderGridRowViewModel> ReaderGridRows { get; } = new();

    public ObservableCollection<PdfPageViewModel> ComparisonPages { get; } = new();

    public ObservableCollection<PdfPageViewModel> ComparisonReaderPages { get; } = new();

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ReaderLayout))]
    [NotifyPropertyChangedFor(nameof(IsSinglePageReaderLayout))]
    [NotifyPropertyChangedFor(nameof(IsContinuousReaderLayout))]
    [NotifyPropertyChangedFor(nameof(IsTwoPageReaderLayout))]
    [NotifyPropertyChangedFor(nameof(IsGridReaderLayout))]
    private ReaderLayoutChoice _selectedReaderLayoutChoice =
        new(ReaderLayoutMode.Continuous, "Continuous", "Scroll through pages vertically.");

    [ObservableProperty]
    private bool _isPageNightMode;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ComparisonPagePosition))]
    private bool _isComparisonOpen;

    [ObservableProperty]
    private string _comparisonDocumentName = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ComparisonPagePosition))]
    private PdfPageViewModel? _comparisonSelectedPage;

    public ReaderLayoutMode ReaderLayout => SelectedReaderLayoutChoice.Mode;

    public bool IsSinglePageReaderLayout => ReaderLayout == ReaderLayoutMode.SinglePage;

    public bool IsContinuousReaderLayout => ReaderLayout == ReaderLayoutMode.Continuous;

    public bool IsTwoPageReaderLayout => ReaderLayout == ReaderLayoutMode.TwoPage;

    public bool IsGridReaderLayout => ReaderLayout == ReaderLayoutMode.Grid;

    public ReaderGridRowViewModel? SelectedReaderGridRow =>
        ReaderGridRows.FirstOrDefault(row => row.Contains(SelectedPage));

    public int PrimaryReaderColumnSpan => IsComparisonOpen ? 1 : 3;

    public string ComparisonPagePosition => ComparisonSelectedPage is null
        ? "No page"
        : $"Page {ComparisonSelectedPage.PageNumber} of {ComparisonPages.Count}";

    partial void OnSelectedReaderLayoutChoiceChanged(ReaderLayoutChoice value) {
        RefreshReaderPages();
        _zoomMode = value.Mode switch {
            ReaderLayoutMode.Continuous => ViewerZoomMode.FitWidth,
            ReaderLayoutMode.Grid => ViewerZoomMode.Grid,
            _ => ViewerZoomMode.FitPage
        };
        ApplyFitZoom();
    }

    partial void OnIsPageNightModeChanged(bool value) {
        foreach (PdfPageViewModel page in Pages) page.IsNightMode = value;
        foreach (PdfPageViewModel page in ComparisonPages) page.IsNightMode = value;
    }

    partial void OnIsComparisonOpenChanged(bool value) => OnPropertyChanged(nameof(PrimaryReaderColumnSpan));

    partial void OnComparisonSelectedPageChanged(PdfPageViewModel? value) {
        OnPropertyChanged(nameof(ComparisonPagePosition));
        ComparisonReaderPages.Clear();
        if (value is not null) ComparisonReaderPages.Add(value);
        OnPropertyChanged(nameof(ComparisonSelectedPage));
        if (_synchronizingComparison || value is null || Pages.Count == 0) return;
        _synchronizingComparison = true;
        try {
            SelectedPage = Pages[Math.Clamp(value.PageNumber, 1, Pages.Count) - 1];
        } finally {
            _synchronizingComparison = false;
        }
    }

    [RelayCommand]
    private void TogglePageNightMode() => IsPageNightMode = !IsPageNightMode;

    [RelayCommand]
    private async Task OpenComparisonAsync(CancellationToken cancellationToken) {
        if (!HasDocument || IsWorkspaceBusy || IsOpening) return;
        string? path = await _pickPdf(cancellationToken).ConfigureAwait(true);
        if (string.IsNullOrWhiteSpace(path)) return;
        string fullPath = Path.GetFullPath(path);
        if (string.Equals(fullPath, DocumentPath, RecentDocumentPathComparison)) {
            OperationStatus = "Choose a different PDF for side-by-side comparison.";
            return;
        }

        _comparisonCancellation?.Cancel();
        _comparisonCancellation?.Dispose();
        _comparisonCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        CancellationTokenSource currentCancellation = _comparisonCancellation;
        PdfWorkspace? candidateWorkspace = null;
        PageSceneCoordinator? candidateSceneCoordinator = null;
        PageRenderCoordinator? candidateRenderCoordinator = null;
        PdfPageViewModel[] candidatePages = [];
        IsOpening = true;
        ErrorMessage = null;

        try {
            candidateWorkspace = await OpenWorkspaceWithPasswordAsync(fullPath, currentCancellation.Token).ConfigureAwait(true);
            if (candidateWorkspace is null) return;
            PdfDocumentSession session = PdfDocumentSession.FromWorkspace(candidateWorkspace);
            candidateSceneCoordinator = new PageSceneCoordinator(session.LoadPageSceneAsync);
            candidateRenderCoordinator = new PageRenderCoordinator(session.RenderPageAsync);
            candidatePages = session.Pages.Select(page => {
                var viewModel = new PdfPageViewModel(
                    page.PageNumber,
                    page.Width,
                    page.Height,
                    page.RotationDegrees,
                    Zoom,
                    candidateSceneCoordinator,
                    candidateRenderCoordinator) {
                    EditorTool = PdfEditorTool.Select,
                    SelectionMode = PdfEditorSelectionMode.None,
                    IsNightMode = IsPageNightMode
                };
                viewModel.LinkActivated += OnComparisonPageLinkActivated;
                return viewModel;
            }).ToArray();
            currentCancellation.Token.ThrowIfCancellationRequested();

            ReaderLayoutChoice layoutToRestore = _layoutBeforeComparison ?? SelectedReaderLayoutChoice;
            CloseComparisonSession(restoreLayout: false, cancelOpen: false);
            _comparisonSession = session;
            _comparisonSceneCoordinator = candidateSceneCoordinator;
            _comparisonRenderCoordinator = candidateRenderCoordinator;
            candidateSceneCoordinator = null;
            candidateRenderCoordinator = null;
            foreach (PdfPageViewModel page in candidatePages) ComparisonPages.Add(page);
            candidatePages = [];
            ComparisonDocumentName = session.FileName;
            IsComparisonOpen = true;
            _layoutBeforeComparison = layoutToRestore;
            SelectedReaderLayoutChoice = ReaderLayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.SinglePage);
            SynchronizeComparisonToPrimary(SelectedPage);
            OperationStatus = "Side-by-side comparison is synchronized by page and zoom.";
        } catch (OperationCanceledException) when (currentCancellation.IsCancellationRequested) {
            OperationStatus = "Comparison opening cancelled";
        } catch (Exception ex) {
            ErrorMessage = ex.Message;
        } finally {
            foreach (PdfPageViewModel page in candidatePages) page.Dispose();
            candidateSceneCoordinator?.Dispose();
            candidateRenderCoordinator?.Dispose();
            candidateWorkspace?.Dispose();
            if (ReferenceEquals(_comparisonCancellation, currentCancellation)) {
                _comparisonCancellation = null;
                currentCancellation.Dispose();
                IsOpening = false;
            }
        }
    }

    [RelayCommand]
    private void CloseComparison() => CloseComparisonSession(restoreLayout: true);

    [RelayCommand]
    private void FirstPage() {
        if (Pages.Count > 0) SelectedPage = Pages[0];
    }

    [RelayCommand]
    private void LastPage() {
        if (Pages.Count > 0) SelectedPage = Pages[^1];
    }

    private void SynchronizeComparisonToPrimary(PdfPageViewModel? primaryPage) {
        if (_synchronizingComparison || !IsComparisonOpen || primaryPage is null || ComparisonPages.Count == 0) return;
        _synchronizingComparison = true;
        try {
            ComparisonSelectedPage = ComparisonPages[Math.Clamp(primaryPage.PageNumber, 1, ComparisonPages.Count) - 1];
        } finally {
            _synchronizingComparison = false;
        }
    }

    private void CancelComparisonOpen() => _comparisonCancellation?.Cancel();

    private void CloseComparisonSession(bool restoreLayout, bool cancelOpen = true) {
        if (cancelOpen) _comparisonCancellation?.Cancel();
        foreach (PdfPageViewModel page in ComparisonPages) page.Dispose();
        ComparisonPages.Clear();
        ComparisonReaderPages.Clear();
        _comparisonSceneCoordinator?.Dispose();
        _comparisonRenderCoordinator?.Dispose();
        _comparisonSceneCoordinator = null;
        _comparisonRenderCoordinator = null;
        _comparisonSession = null;
        ComparisonSelectedPage = null;
        ComparisonDocumentName = string.Empty;
        IsComparisonOpen = false;
        OnPropertyChanged(nameof(ComparisonPagePosition));

        ReaderLayoutChoice? restore = _layoutBeforeComparison;
        _layoutBeforeComparison = null;
        if (restoreLayout && restore is not null) SelectedReaderLayoutChoice = restore;
    }

    private void RefreshReaderPages() {
        if (ReaderLayout == ReaderLayoutMode.Grid) {
            ReaderPages.Clear();
            RefreshReaderGridRows();
            OnPropertyChanged(nameof(SelectedPage));
            OnPropertyChanged(nameof(SelectedReaderGridRow));
            return;
        }

        ReaderGridRows.Clear();
        IReadOnlyList<PdfPageViewModel> desiredPages = ReaderLayout switch {
            ReaderLayoutMode.SinglePage => SelectedPage is null ? [] : [SelectedPage],
            ReaderLayoutMode.TwoPage => GetSelectedSpread(),
            _ => Pages
        };
        if (!ReaderPages.SequenceEqual(desiredPages)) {
            ReaderPages.Clear();
            foreach (PdfPageViewModel page in desiredPages) ReaderPages.Add(page);
            OnPropertyChanged(nameof(SelectedPage));
        }
        OnPropertyChanged(nameof(SelectedReaderGridRow));
    }

    private void RefreshReaderGridRows() {
        int columns = GetReaderGridColumnCount();
        ReaderGridRowViewModel[] desiredRows = Pages
            .Chunk(columns)
            .Select(static pages => new ReaderGridRowViewModel(pages))
            .ToArray();
        if (ReaderGridRows.Count == desiredRows.Length &&
            ReaderGridRows.Zip(desiredRows).All(pair => pair.First.Pages.SequenceEqual(pair.Second.Pages))) return;

        ReaderGridRows.Clear();
        foreach (ReaderGridRowViewModel row in desiredRows) ReaderGridRows.Add(row);
    }

    private int GetReaderGridColumnCount() => GetGridColumnCount(Math.Max(200D, _viewportWidth - 72D));

    private IReadOnlyList<PdfPageViewModel> GetSelectedSpread() {
        if (Pages.Count == 0) return [];
        int selectedPageNumber = Math.Clamp(SelectedPage?.PageNumber ?? 1, 1, Pages.Count);
        if (selectedPageNumber == 1) return [Pages[0]];

        int firstPageNumber = selectedPageNumber % 2 == 0
            ? selectedPageNumber
            : selectedPageNumber - 1;
        return firstPageNumber < Pages.Count
            ? [Pages[firstPageNumber - 1], Pages[firstPageNumber]]
            : [Pages[firstPageNumber - 1]];
    }
}
