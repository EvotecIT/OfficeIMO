using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using System.ComponentModel;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Features.Mobile;

/// <summary>Touch workspace over the same document state, renderer, and editing commands as desktop Studio.</summary>
public sealed partial class MobileWorkspaceView : UserControl {
    private Control? _sheetOpener;
    private bool _showPages = true;
    private Func<Task>? _shareDocumentAsync;
    private MainWindowViewModel? _observedDocument;
    private bool _submittingNote;
    private Func<Task>? _openSampleAsync;
    private readonly HashSet<MainWindowViewModel> _configuredDocuments = [];
    private readonly Dictionary<MainWindowViewModel, string> _noteDrafts = [];

    internal Func<Task>? OpenSampleAsync {
        get => _openSampleAsync;
        set { _openSampleAsync = value; SampleButton.IsVisible = value is not null; }
    }

    public MobileWorkspaceView() {
        InitializeComponent();
        InitializeTouchInput();
        InitializeNavigation();
        SizeChanged += (_, _) => UpdateLayoutMode();
        PageScroll.SizeChanged += (_, e) => Document?.SetViewportSize(e.NewSize.Width, e.NewSize.Height);
        DataContextChanged += (_, _) => ObserveDocument();
        AttachedToVisualTree += (_, _) => ObserveDocument();
        DetachedFromVisualTree += (_, _) => {
            if (_observedDocument is not null) _observedDocument.PropertyChanged -= OnDocumentChanged;
            _observedDocument = null;
        };
        KeyDown += (_, e) => {
            if (e.Key == Key.Escape && SheetScrim.IsVisible) { DismissSheet(); e.Handled = true; }
            else if (e.Key == Key.Escape && SearchPanel.IsVisible) { CloseSearch(); e.Handled = true; }
            else if (!SheetScrim.IsVisible && e.KeyModifiers.HasFlag(KeyModifiers.Meta)) {
                if (e.Key == Key.O) { Document?.OpenCommand.Execute(null); e.Handled = true; }
                else if (e.Key == Key.W && _controller is not null) { _ = _controller.Tabs.CloseSelectedTabAsync(); e.Handled = true; }
                else if (e.Key == Key.F) { OnSearchClick(SearchButton, new RoutedEventArgs()); e.Handled = true; }
                else if (e.Key is Key.OemOpenBrackets or Key.OemCloseBrackets && e.KeyModifiers.HasFlag(KeyModifiers.Shift)) {
                    _controller?.Tabs.SelectRelativeTab(e.Key == Key.OemOpenBrackets); e.Handled = true;
                }
            }
        };
    }

    internal MainWindowViewModel? Document => DataContext as MainWindowViewModel;

    private void ObserveDocument() {
        if (_observedDocument is not null) {
            _observedDocument.PropertyChanged -= OnDocumentChanged;
            _noteDrafts[_observedDocument] = NoteText.Text ?? string.Empty;
        }
        _observedDocument = Document;
        NoteText.Text = Document is { } active && _noteDrafts.TryGetValue(active, out string? draft) ? draft : string.Empty;
        _pinchStartZoom = null;
        PageScroll.Offset = default;
        RefreshPageList();
        Document?.SetViewportSize(PageScroll.Bounds.Width, PageScroll.Bounds.Height);
        if (_observedDocument is null) return;
        _observedDocument.PropertyChanged += OnDocumentChanged;
        if (_configuredDocuments.Add(_observedDocument)) FitNewDocument();
    }

    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs e) {
        // Opening restores desktop view preferences; the mobile surface always presents one page at a time.
        if (e.PropertyName == nameof(MainWindowViewModel.IsOpening) && Document?.IsOpening == false) {
            FitNewDocument();
            RefreshPageList();
        }
        if (e.PropertyName == nameof(MainWindowViewModel.SelectedPage)) SelectCurrentThumbnail();
    }

    private void FitNewDocument() {
        if (Document is not { } document) return;
        document.SelectedReaderLayoutChoice = document.ReaderLayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.SinglePage);
        document.FitPageCommand.Execute(null);
    }

    /// <summary>The host saves a current copy and presents its platform share surface.</summary>
    internal Func<Task>? ShareDocumentAsync {
        get => _shareDocumentAsync;
        set { _shareDocumentAsync = value; ShareButton.IsVisible = value is not null; }
    }

    private bool HasSidebarRoom => Bounds.Width >= 720 && Bounds.Height >= 500;

    private void UpdateLayoutMode() {
        bool wide = Bounds.Width >= 720;
        bool sidebar = HasSidebarRoom && _showPages && Document?.HasDocument == true;
        WideTools.IsVisible = wide;
        CompactTools.IsVisible = !wide;
        WideFit.IsVisible = wide;
        ShareLabel.IsVisible = wide;
        PageSidebar.IsVisible = sidebar;
        WorkspaceGrid.ColumnDefinitions[0].Width = new GridLength(sidebar ? 196 : 0);
        SheetCard.MaxHeight = Math.Max(180, Bounds.Height - 32);
        SheetCard.Margin = wide ? new Thickness(20) : default;
        SheetCard.CornerRadius = wide ? new CornerRadius(22) : new CornerRadius(22, 22, 0, 0);
        if (HasSidebarRoom && SheetScrim.IsVisible && SheetPages.IsVisible) DismissSheet();
        bool pageSheet = SheetScrim.IsVisible && SheetPages.IsVisible;
        HostPageList(pageSheet ? SheetPages : sidebar ? SidebarPages : null);
        _pageList.Height = pageSheet ? Math.Min(420, Math.Max(120, Bounds.Height - 180)) : double.NaN;
    }

    private void OnPagesClick(object? sender, RoutedEventArgs e) {
        if (HasSidebarRoom) { _showPages = !_showPages; UpdateLayoutMode(); }
        else {
            ShowSheet("Pages", SheetPages, sender as Control);
            UpdateLayoutMode();
            _pageList.ScrollIntoView(_pageList.SelectedItem!);
        }
    }

    private void OnSearchClick(object? sender, RoutedEventArgs e) {
        SearchPanel.IsVisible = !SearchPanel.IsVisible;
        if (SearchPanel.IsVisible) MobileSearchBox.Focus();
        else PageScroll.Focus();
    }

    private void OnSearchKeyDown(object? sender, KeyEventArgs e) {
        if (e.Key == Key.Enter) { Document?.SearchCommand.Execute(null); e.Handled = true; }
    }

    private void OnCloseSearchClick(object? sender, RoutedEventArgs e) => CloseSearch();
    private void CloseSearch() { SearchPanel.IsVisible = false; PagesButton.Focus(); }

    private void OnNoteClick(object? sender, RoutedEventArgs e) {
        if (Document?.CanEditAnnotations != true) {
            if (Document is { } document) document.ErrorMessage = "Annotations are unavailable for this document.";
            return;
        }
        ShowSheet("New note", NotePanel, sender as Control);
        NoteText.Focus();
    }

    private void ShowSheet(string title, Control content, Control? opener) {
        _sheetOpener = opener;
        SheetTitle.Text = title;
        SheetPages.IsVisible = ReferenceEquals(content, SheetPages);
        NoteScroll.IsVisible = ReferenceEquals(content, NotePanel);
        SheetScrim.IsVisible = true;
        HeaderBar.IsEnabled = TabBar.IsEnabled = SearchPanel.IsEnabled = WorkspaceGrid.IsEnabled = FooterBar.IsEnabled = false;
        SheetDone.Focus();
    }

    private void OnDismissClick(object? sender, RoutedEventArgs e) => DismissSheet();

    private void DismissSheet() {
        if (_submittingNote) return;
        SheetScrim.IsVisible = false;
        HeaderBar.IsEnabled = TabBar.IsEnabled = SearchPanel.IsEnabled = WorkspaceGrid.IsEnabled = FooterBar.IsEnabled = true;
        UpdateLayoutMode();
        _sheetOpener?.Focus();
    }

    private async void OnAddNoteClick(object? sender, RoutedEventArgs e) {
        if (_submittingNote || Document is not { SelectedPage: { } page } document || document.IsWorkspaceBusy || string.IsNullOrWhiteSpace(NoteText.Text)) return;
        _submittingNote = true;
        NotePanel.IsEnabled = SheetDone.IsEnabled = false;
        try {
            document.EditorText = NoteText.Text.Trim();
            await document.ApplyPageMarkupAsync(PdfEditorTool.Note, new PdfEditorGesture(page.PageNumber, 24, 24, 48, 48, []));
            if (document.HasError || !ReferenceEquals(document, Document)) return;
            document.ShowViewModeCommand.Execute(null);
            NoteText.Text = string.Empty;
            _submittingNote = false;
            DismissSheet();
        } catch (Exception error) {
            document.ErrorMessage = error.Message;
        } finally {
            _submittingNote = false;
            NotePanel.IsEnabled = SheetDone.IsEnabled = true;
        }
    }

    private async void OnShareClick(object? sender, RoutedEventArgs e) {
        if (ShareDocumentAsync is null || Document?.IsWorkspaceBusy == true) return;
        try { await ShareDocumentAsync(); }
        catch (Exception error) { if (Document is { } document) document.ErrorMessage = error.Message; }
    }

    private async void OnSampleClick(object? sender, RoutedEventArgs e) {
        if (_openSampleAsync is null || !SampleButton.IsEnabled) return;
        SampleButton.IsEnabled = false;
        try { await _openSampleAsync(); }
        catch (Exception error) { if (Document is { } document) document.ErrorMessage = error.Message; }
        finally { SampleButton.IsEnabled = true; }
    }
}
