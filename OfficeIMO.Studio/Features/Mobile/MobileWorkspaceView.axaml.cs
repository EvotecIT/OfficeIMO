using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Threading;
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
            CompleteCloseDecision(UnsavedChangesDecision.Cancel);
            if (_observedDocument is not null) _observedDocument.PropertyChanged -= OnDocumentChanged;
            _observedDocument = null;
        };
        KeyDown += OnWorkspaceKeyDown;
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
        UpdateDocumentStatus();
        if (_configuredDocuments.Add(_observedDocument)) FitNewDocument();
    }

    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs e) {
        // Opening restores desktop view preferences; the mobile surface always presents one page at a time.
        if (e.PropertyName == nameof(MainWindowViewModel.IsOpening) && Document?.IsOpening == false) {
            FitNewDocument();
            RefreshPageList();
        }
        if (e.PropertyName == nameof(MainWindowViewModel.SelectedPage)) SelectCurrentThumbnail();
        if (e.PropertyName is nameof(MainWindowViewModel.IsDirty) or nameof(MainWindowViewModel.IsWorkspaceBusy) or
            nameof(MainWindowViewModel.HasDocument) or nameof(MainWindowViewModel.IsOpening)) UpdateDocumentStatus();
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
        CompactTools.IsVisible = !wide && Document?.HasDocument == true;
        WideFit.IsVisible = wide;
        ShareLabel.IsVisible = wide;
        PageSidebar.IsVisible = sidebar;
        WorkspaceGrid.ColumnDefinitions[0].Width = new GridLength(sidebar ? 196 : 0);
        SheetCard.MaxHeight = Math.Max(180, Bounds.Height - 32);
        SheetCard.Margin = wide ? new Thickness(20) : default;
        SheetCard.VerticalAlignment = wide ? VerticalAlignment.Center : VerticalAlignment.Bottom;
        SheetCard.CornerRadius = wide ? new CornerRadius(22) : new CornerRadius(22, 22, 0, 0);
        if (HasSidebarRoom && SheetScrim.IsVisible && SheetPages.IsVisible) DismissSheet();
        bool pageSheet = SheetScrim.IsVisible && SheetPages.IsVisible;
        HostPageList(pageSheet ? SheetPages : sidebar ? SidebarPages : null);
        _pageList.Height = pageSheet ? Math.Min(420, Math.Max(120, Bounds.Height - 180)) : double.NaN;
        DocumentList.MaxHeight = Math.Max(100, Bounds.Height - 180);
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
        if (SearchPanel.IsVisible) CloseSearch();
        else OpenSearch();
    }

    private void OnSearchKeyDown(object? sender, KeyEventArgs e) {
        if (e.Key != Key.Enter || Document is not { } document) return;
        if (!document.HasSearchResults) ExecuteIfAvailable(document.SearchCommand);
        else if (e.KeyModifiers.HasFlag(KeyModifiers.Shift)) document.PreviousSearchResultCommand.Execute(null);
        else document.NextSearchResultCommand.Execute(null);
        e.Handled = true;
    }

    private void OnCloseSearchClick(object? sender, RoutedEventArgs e) => CloseSearch();
    private void OpenSearch() { SearchPanel.IsVisible = true; MobileSearchBox.Focus(); }
    private void CloseSearch() {
        Document?.ClearSearchCommand.Execute(null);
        SearchPanel.IsVisible = false;
        PageScroll.Focus();
    }

    private void OnNoteClick(object? sender, RoutedEventArgs e) {
        if (Document?.CanEditAnnotations != true) {
            if (Document is { } document) document.ErrorMessage = "Annotations are unavailable for this document.";
            return;
        }
        ShowSheet("New note", NotePanel, sender as Control);
    }

    private void ShowSheet(string title, Control content, Control? opener) {
        _sheetOpener = opener;
        SheetTitle.Text = title;
        SheetPages.IsVisible = ReferenceEquals(content, SheetPages);
        NoteScroll.IsVisible = ReferenceEquals(content, NotePanel);
        DocumentList.IsVisible = ReferenceEquals(content, DocumentList);
        CloseScroll.IsVisible = ReferenceEquals(content, CloseScroll);
        SheetDone.IsVisible = !CloseScroll.IsVisible;
        SheetScrim.IsVisible = true;
        HeaderBar.IsEnabled = TabBar.IsEnabled = SearchPanel.IsEnabled = WorkspaceGrid.IsEnabled = FooterBar.IsEnabled = false;
        Control focus = CloseScroll.IsVisible ? CloseCancel : NoteScroll.IsVisible ? NoteText : DocumentList.IsVisible ? DocumentList : SheetDone;
        focus.Focus();
        // ScrollViewer content can join the visual tree only after its first visible layout.
        Dispatcher.UIThread.Post(() => {
            if (SheetScrim.IsVisible && content.IsEffectivelyVisible) focus.Focus();
        }, DispatcherPriority.Loaded);
    }

    private void OnDismissClick(object? sender, RoutedEventArgs e) => DismissSheet();

    private void DismissSheet() {
        if (_submittingNote) return;
        if (_closeDecision is not null) { CompleteCloseDecision(UnsavedChangesDecision.Cancel); return; }
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
