using System.Collections.ObjectModel;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>An independently navigable viewport over a tab's canonical document session.</summary>
public sealed partial class StudioDocumentPaneViewModel : ObservableObject, IDisposable {
    private readonly StudioDocumentPanes _host;
    private readonly List<PdfPageViewModel> _sourcePages = [];
    private bool _synchronizing;
    private bool _applyingFit;
    private bool _disposed;
    private double _width = 500, _height = 500;
    private ViewerZoomMode _zoomMode = ViewerZoomMode.Custom;

    internal StudioDocumentPaneViewModel(StudioDocumentPanes host, StudioDocumentTabViewModel tab) {
        _host = host;
        Tab = tab;
        _zoom = Document.Zoom;
        _zoomMode = Document.ReaderZoomMode;
        _selectedLayout = Document.SelectedReaderLayoutChoice;
        Document.PropertyChanged += OnDocumentChanged;
        Document.ReaderPresentationReplaced += OnPresentationReplaced;
        RebuildPages(Document.SelectedPage?.PageNumber ?? 1);
    }

    internal StudioDocumentTabViewModel Tab { get; }
    internal MainWindowViewModel Document => Tab.Document;
    public StudioDocumentTabViewModel SelectedTab { get => Tab; set { if (value is not null && value != Tab) _host.AssignDocument(this, value); } }
    public ObservableCollection<StudioDocumentTabViewModel> OpenDocuments => _host.Tabs;
    public string Title => Document.DocumentName;
    public string PagePosition => Document.PaneLocalizer.Format("Document.PagePosition", SelectedPage?.PageNumber ?? 0, Pages.Count);
    public string ZoomLabel => $"{Zoom:P0}";
    public string CloseLabel => Document.PaneLocalizer.GetOrDefault("Panes.Close", "Close pane");
    public string DocumentLabel => Document.PaneLocalizer.GetOrDefault("Panes.Document", "Document in this pane");
    public string ActivateLabel => Document.PaneLocalizer.GetOrDefault("Panes.Activate", "Activate pane");
    public IReadOnlyList<ReaderLayoutChoice> LayoutChoices => Document.ReaderLayoutChoices;
    public ObservableCollection<PdfPageViewModel> Pages { get; } = new();
    public ObservableCollection<PdfPageViewModel> ReaderPages { get; } = new();
    public ObservableCollection<ReaderGridRowViewModel> GridRows { get; } = new();
    public ReaderGridRowViewModel? SelectedGridRow => GridRows.FirstOrDefault(row => row.Contains(SelectedPage));
    public bool IsGrid => SelectedLayout.Mode == ReaderLayoutMode.Grid;
    public bool IsContinuous => SelectedLayout.Mode == ReaderLayoutMode.Continuous;
    public bool IsTwoPage => SelectedLayout.Mode == ReaderLayoutMode.TwoPage;
    public bool CanPrevious => SelectedPage?.PageNumber > 1;
    public bool CanNext => SelectedPage is { } page && page.PageNumber < Pages.Count;

    [ObservableProperty] private bool _isActive;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(PagePosition))] [NotifyPropertyChangedFor(nameof(CanPrevious))]
    [NotifyPropertyChangedFor(nameof(CanNext))] [NotifyPropertyChangedFor(nameof(SelectedGridRow))]
    private PdfPageViewModel? _selectedPage;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(ZoomLabel))] private double _zoom;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(IsGrid))] [NotifyPropertyChangedFor(nameof(IsContinuous))]
    [NotifyPropertyChangedFor(nameof(IsTwoPage))] private ReaderLayoutChoice _selectedLayout;

    partial void OnSelectedPageChanged(PdfPageViewModel? value) {
        RefreshReaderPages();
        if (_zoomMode != ViewerZoomMode.Custom) Fit();
        PublishNavigation();
    }
    partial void OnZoomChanged(double value) {
        if (!_synchronizing && !_applyingFit) _zoomMode = ViewerZoomMode.Custom;
        foreach (var page in Pages) page.SetZoom(value);
        PublishNavigation();
    }
    partial void OnSelectedLayoutChanged(ReaderLayoutChoice value) {
        RefreshReaderPages();
        if (!_synchronizing) {
            _zoomMode = value.Mode is ReaderLayoutMode.Continuous ? ViewerZoomMode.FitWidth :
                value.Mode is ReaderLayoutMode.Grid ? ViewerZoomMode.Grid : ViewerZoomMode.FitPage;
            Fit();
        }
        PublishNavigation();
    }
    partial void OnIsActiveChanged(bool value) => MirrorInteractions();

    internal void Activate() => _host.Activate(this);
    internal void ActivatePage(int pageNumber) {
        Activate();
        SelectedPage = Pages.ElementAtOrDefault(pageNumber - 1);
        PublishNavigation();
    }
    internal void PublishNavigation() {
        if (!IsActive || _synchronizing || _disposed || Document.IsReplacingReaderPresentation) return;
        _synchronizing = true;
        try { Document.ApplyPaneNavigation(SelectedPage?.PageNumber ?? 1, Zoom, SelectedLayout.Mode, _width, _height); }
        finally { _synchronizing = false; }
    }
    internal void SetViewport(double width, double height) {
        if (width <= 0 || height <= 0) return;
        _width = width; _height = height;
        if (IsGrid) RefreshReaderPages();
        if (_zoomMode != ViewerZoomMode.Custom) Fit();
    }

    [RelayCommand] private void Previous() { Activate(); if (CanPrevious) SelectedPage = Pages[SelectedPage!.PageNumber - 2]; }
    [RelayCommand] private void Next() { Activate(); if (CanNext) SelectedPage = Pages[SelectedPage!.PageNumber]; }
    [RelayCommand] private void ZoomIn() { Activate(); _zoomMode = ViewerZoomMode.Custom; Zoom = Math.Round(Math.Min(3, Zoom + .25), 2); }
    [RelayCommand] private void ZoomOut() { Activate(); _zoomMode = ViewerZoomMode.Custom; Zoom = Math.Round(Math.Max(.25, Zoom - .25), 2); }
    [RelayCommand] private void FitWidth() { Activate(); _zoomMode = ViewerZoomMode.FitWidth; Fit(); }
    [RelayCommand] private void FitPage() { Activate(); _zoomMode = ViewerZoomMode.FitPage; Fit(); }
    [RelayCommand] private void Close() => _host.ClosePane(this);
    [RelayCommand] private void ActivatePane() => Activate();

    internal bool NavigateTypedPage(string? text) {
        if (!PdfPageNavigation.TryParsePageNumber(text, Pages.Count, out int page)) return false;
        ActivatePage(page); return true;
    }

    private void Fit() {
        if (SelectedPage is null) return;
        double width = SelectedPage.DisplayWidth / Math.Max(Zoom, .01);
        double height = SelectedPage.DisplayHeight / Math.Max(Zoom, .01);
        double availableWidth = Math.Max(120, _width - 48), availableHeight = Math.Max(120, _height - 48);
        double fitted = _zoomMode switch {
            ViewerZoomMode.FitWidth => Math.Min(2, Math.Min(880, availableWidth) / width),
            ViewerZoomMode.Grid => Math.Max(120, (availableWidth - 24 * (GridColumns + 1)) / GridColumns) / width,
            _ => Math.Min(availableWidth / (width * (IsTwoPage && ReaderPages.Count > 1 ? 2 : 1)), availableHeight / height)
        };
        _applyingFit = true;
        try { Zoom = Math.Round(Math.Clamp(fitted, .25, 3), 2); }
        finally { _applyingFit = false; }
    }
    private int GridColumns => PdfReaderViewportLayout.GridColumnCount(Math.Max(120, _width - 48));
    private void RefreshReaderPages() {
        ReaderPages.Clear(); GridRows.Clear();
        if (IsGrid) {
            foreach (var pages in Pages.Chunk(GridColumns)) GridRows.Add(new(pages));
        } else {
            IEnumerable<PdfPageViewModel> visible = SelectedLayout.Mode switch {
                ReaderLayoutMode.SinglePage => SelectedPage is null ? [] : [SelectedPage],
                ReaderLayoutMode.TwoPage => SelectedSpread(),
                _ => Pages
            };
            foreach (var page in visible) ReaderPages.Add(page);
        }
        OnPropertyChanged(nameof(SelectedGridRow));
    }
    private IEnumerable<PdfPageViewModel> SelectedSpread() {
        var spread = PdfReaderViewportLayout.Spread(SelectedPage?.PageNumber ?? 1, Pages.Count);
        return Pages.Skip(spread.StartIndex).Take(spread.Count);
    }
    private void OnPresentationReplaced(object? sender, EventArgs args) => RebuildPages(SelectedPage?.PageNumber ?? 1);
    private void RebuildPages(int selectedPage) {
        _synchronizing = true;
        try {
            foreach (var page in _sourcePages) page.PropertyChanged -= OnSourcePageChanged;
            _sourcePages.Clear();
            foreach (var page in Pages) page.Dispose();
            Pages.Clear();
            foreach (var page in Document.CreatePanePages(ActivatePage)) Pages.Add(page);
            foreach (var source in Document.Pages) { _sourcePages.Add(source); source.PropertyChanged += OnSourcePageChanged; }
            SelectedPage = Pages.ElementAtOrDefault(Math.Clamp(selectedPage - 1, 0, Math.Max(0, Pages.Count - 1)));
            foreach (var page in Pages) page.SetZoom(Zoom);
            MirrorInteractions(); RefreshReaderPages();
            OnPropertyChanged(nameof(PagePosition));
        } finally { _synchronizing = false; }
    }
    private void OnSourcePageChanged(object? sender, PropertyChangedEventArgs args) {
        if (sender is PdfPageViewModel source && Pages.ElementAtOrDefault(source.PageNumber - 1) is { } page)
            page.MirrorInteractionState(source, IsActive);
    }
    private void MirrorInteractions() {
        for (int index = 0; index < Math.Min(Pages.Count, Document.Pages.Count); index++)
            Pages[index].MirrorInteractionState(Document.Pages[index], IsActive);
    }
    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName == nameof(MainWindowViewModel.DocumentName)) OnPropertyChanged(nameof(Title));
        if (args.PropertyName == nameof(MainWindowViewModel.IsComparisonOpen) && !Document.IsComparisonOpen) PublishNavigation();
        if (_synchronizing || !IsActive || Document.IsReplacingReaderPresentation || Document.IsComparisonOpen) return;
        _synchronizing = true;
        try {
            if (args.PropertyName == nameof(MainWindowViewModel.SelectedPage))
                SelectedPage = Pages.ElementAtOrDefault((Document.SelectedPage?.PageNumber ?? 0) - 1);
            else if (args.PropertyName == nameof(MainWindowViewModel.Zoom)) { _zoomMode = Document.ReaderZoomMode; Zoom = Document.Zoom; }
            else if (args.PropertyName == nameof(MainWindowViewModel.SelectedReaderLayoutChoice)) {
                _zoomMode = Document.ReaderZoomMode;
                SelectedLayout = Document.SelectedReaderLayoutChoice;
            }
        } finally { _synchronizing = false; }
    }
    public void Dispose() {
        if (_disposed) return; _disposed = true;
        Document.PropertyChanged -= OnDocumentChanged;
        Document.ReaderPresentationReplaced -= OnPresentationReplaced;
        foreach (var page in _sourcePages) page.PropertyChanged -= OnSourcePageChanged;
        foreach (var page in Pages) page.Dispose();
        Pages.Clear(); ReaderPages.Clear(); GridRows.Clear();
    }
}
