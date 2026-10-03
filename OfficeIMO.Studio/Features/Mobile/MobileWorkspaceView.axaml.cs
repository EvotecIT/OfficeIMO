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

    internal Func<Task>? OpenSampleAsync {
        get => _openSampleAsync;
        set { _openSampleAsync = value; SampleButton.IsVisible = value is not null; }
    }

    public MobileWorkspaceView() {
        InitializeComponent();
        InitializeTouchInput();
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
        };
    }

    internal MainWindowViewModel? Document => DataContext as MainWindowViewModel;

    private void ObserveDocument() {
        if (_observedDocument is not null) _observedDocument.PropertyChanged -= OnDocumentChanged;
        _observedDocument = Document;
        if (_observedDocument is null) return;
        _observedDocument.PropertyChanged += OnDocumentChanged;
        FitNewDocument();
    }

    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs e) {
        // Opening restores desktop view preferences; the mobile surface always presents one page at a time.
        if (e.PropertyName == nameof(MainWindowViewModel.IsOpening) && Document?.IsOpening == false) FitNewDocument();
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

    private void UpdateLayoutMode() {
        bool sidebar = Bounds.Width >= 900 && _showPages;
        PageSidebar.IsVisible = sidebar;
        WorkspaceGrid.ColumnDefinitions[0].Width = new GridLength(sidebar ? 220 : 0);
    }

    private void OnPagesClick(object? sender, RoutedEventArgs e) {
        if (Bounds.Width >= 900) { _showPages = !_showPages; UpdateLayoutMode(); }
        else ShowSheet("Pages", SheetPages, sender as Control);
    }

    private void OnSearchClick(object? sender, RoutedEventArgs e) {
        ShowSheet("Find in document", SearchPanel, sender as Control);
        MobileSearchBox.Focus();
    }

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
        SearchPanel.IsVisible = ReferenceEquals(content, SearchPanel);
        NotePanel.IsVisible = ReferenceEquals(content, NotePanel);
        SheetScrim.IsVisible = true;
        HeaderBar.IsEnabled = WorkspaceGrid.IsEnabled = FooterBar.IsEnabled = false;
        SheetDone.Focus();
    }

    private void OnDismissClick(object? sender, RoutedEventArgs e) => DismissSheet();

    private void DismissSheet() {
        if (_submittingNote) return;
        SheetScrim.IsVisible = false;
        HeaderBar.IsEnabled = WorkspaceGrid.IsEnabled = FooterBar.IsEnabled = true;
        _sheetOpener?.Focus();
    }

    private void OnSheetPageSelected(object? sender, SelectionChangedEventArgs e) {
        if (SheetPages.IsVisible && SheetScrim.IsVisible) DismissSheet();
        PageScroll.Offset = default;
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
