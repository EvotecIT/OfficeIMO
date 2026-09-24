using Avalonia.Controls;
using Avalonia.Input;
using System.ComponentModel;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Workspace;

public sealed partial class DocumentWorkspaceView : UserControl {
    private bool? _compactLayout;
    private MainWindowViewModel? _document;
    private double _navigationWidth = 238D;
    private double _inspectorWidth = 300D;

    public DocumentWorkspaceView() {
        InitializeComponent();
        SizeChanged += (_, e) => ApplyResponsiveLayout(e.NewSize.Width);
        DataContextChanged += (_, _) => {
            if (_document is not null) {
                CapturePaneWidths();
                _document.PropertyChanged -= OnDocumentChanged;
            }
            _document = DataContext as MainWindowViewModel;
            if (_document is not null) _document.PropertyChanged += OnDocumentChanged;
            RestorePaneWidths();
        };
        NavigationSplitter.AddHandler(PointerReleasedEvent, (_, _) => CapturePaneWidths(), handledEventsToo: true);
        InspectorSplitter.AddHandler(PointerReleasedEvent, (_, _) => CapturePaneWidths(), handledEventsToo: true);
        NavigationSplitter.KeyUp += (_, _) => CapturePaneWidths();
        InspectorSplitter.KeyUp += (_, _) => CapturePaneWidths();
    }

    internal void ApplyResponsiveLayout(double width) {
        CommandRow.Classes.Set("compactCommands", width < 1320D);
        bool compact = width < 1100D;
        if (_compactLayout == compact) return;
        _compactLayout = compact;
        DocumentModeButtons.IsVisible = !compact;
        DocumentModePicker.IsVisible = compact;
        ShowContextPanes();
    }

    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName == nameof(MainWindowViewModel.DocumentViewState)) RestorePaneWidths();
        else if (e.PropertyName is nameof(MainWindowViewModel.DocumentMode) or nameof(MainWindowViewModel.IsFocusReading) or
                 nameof(MainWindowViewModel.HasDocument) or nameof(MainWindowViewModel.IsComparisonOpen)) ShowContextPanes();
        else if (e.PropertyName == nameof(MainWindowViewModel.SelectedObject) && _document?.SelectedObject is not null)
            SetPanes(_compactLayout != true && NavigationPane.IsVisible, true);
        else if (e.PropertyName is nameof(MainWindowViewModel.HasOrganizerSelection)) UpdateOrganizerActionBar();
        if (e.PropertyName is nameof(MainWindowViewModel.SelectedPage) or nameof(MainWindowViewModel.SelectedPagePosition) or
            nameof(MainWindowViewModel.HasDocument)) UpdatePageNumber();
        if (e.PropertyName is nameof(MainWindowViewModel.DocumentMode) or nameof(MainWindowViewModel.EditorInstruction) or
            nameof(MainWindowViewModel.ReaderHint) or nameof(MainWindowViewModel.OrganizerSelectionLabel)) UpdateStatusHint();
    }

    private bool IsPagesGrid => _document is { DocumentMode: StudioDocumentMode.Pages, IsFocusReading: false, HasDocument: true };

    private void UpdateOrganizerActionBar() =>
        OrganizerActionBar.IsVisible = IsPagesGrid && _document?.HasOrganizerSelection == true;

    private void UpdatePageNumber() {
        if (_document is null) return;
        if (!PageNumberBox.IsFocused)
            PageNumberBox.Text = _document.SelectedPage?.PageNumber.ToString(System.Globalization.CultureInfo.CurrentCulture) ?? string.Empty;
        PageCountText.Text = StudioLocalization.Current.Format("Shell.PageCount", _document.Pages.Count);
    }

    private void UpdateStatusHint() {
        if (_document is null) { StatusHint.Text = null; return; }
        StatusHint.Text = _document.DocumentMode switch {
            StudioDocumentMode.View => _document.ReaderHint,
            StudioDocumentMode.Annotate or StudioDocumentMode.Edit => _document.EditorInstruction,
            StudioDocumentMode.Pages => StudioLocalization.Current.Get("Shell.OrganizerStatus"),
            _ => null
        };
    }

    private void OnPageNumberKeyDown(object? sender, KeyEventArgs e) {
        if (e.Key == Key.Enter) {
            GoToTypedPage();
            (GridPagesList.IsEffectivelyVisible ? GridPagesList : PagesList).Focus();
            e.Handled = true;
        } else if (e.Key == Key.Escape) {
            UpdatePageNumber();
            PagesList.Focus();
            e.Handled = true;
        }
    }

    private void OnPageNumberLostFocus(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => UpdatePageNumber();

    private void GoToTypedPage() {
        if (_document is null) return;
        if (int.TryParse(PageNumberBox.Text, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.CurrentCulture, out int page) &&
            page >= 1 && page <= _document.Pages.Count) {
            _document.NavigateToOrganizerPage(page);
        }
        UpdatePageNumber();
    }

    private void OnFindClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => FocusSearch();

    private void ShowContextPanes() {
        UpdatePageNumber();
        UpdateStatusHint();
        if (_document?.IsFocusReading == true) { SetPanes(false, false); return; }
        // Comparing needs the width for two documents; the properties pane returns when the comparison closes.
        if (_document?.IsComparisonOpen == true) { SetPanes(NavigationPane.IsVisible && _compactLayout != true, false); return; }
        StudioDocumentMode mode = _document?.DocumentMode ?? StudioDocumentMode.View;
        // Reading keeps the document details open on wide windows so switching to a task mode does not refit the page.
        bool inspector = mode is not StudioDocumentMode.Pages && (mode != StudioDocumentMode.View || _compactLayout != true);
        bool navigation = mode == StudioDocumentMode.Pages || _compactLayout != true;
        if (_document?.DocumentViewState.Panes.TryGetValue(mode, out StudioPanePreference? preference) == true) {
            navigation = preference.Navigation;
            inspector = preference.Inspector;
        }
        SetPanes(navigation && (!inspector || _compactLayout != true), inspector);
    }

    private void CapturePaneWidths() {
        if (NavigationPane.IsVisible && PageViewport.ColumnDefinitions[0].ActualWidth > 0)
            _navigationWidth = Math.Clamp(PageViewport.ColumnDefinitions[0].ActualWidth, 200D, 320D);
        if (InspectorPane.IsVisible && PageViewport.ColumnDefinitions[2].ActualWidth > 0)
            _inspectorWidth = Math.Clamp(PageViewport.ColumnDefinitions[2].ActualWidth, 280D, 380D);
        _document?.UpdatePanePreferences(_navigationWidth, _inspectorWidth);
    }

    private void RestorePaneWidths() {
        _navigationWidth = _document?.DocumentViewState.NavigationWidth ?? 238D;
        _inspectorWidth = _document?.DocumentViewState.InspectorWidth ?? 300D;
        ShowContextPanes();
    }

    private void SetPanes(bool navigation, bool inspector) {
        bool pagesGrid = IsPagesGrid;
        if (pagesGrid) {
            navigation = true;
            if (NavigationTabs.SelectedIndex != 0) NavigationTabs.SelectedIndex = 0;
        }
        Grid.SetColumnSpan(NavigationPane, pagesGrid ? 2 : 1);
        ReaderCanvas.IsVisible = !pagesGrid && _document?.HasDocument == true;
        NavigationToggle.IsEnabled = !pagesGrid;
        UpdateOrganizerActionBar();
        NavigationPane.IsVisible = navigation;
        InspectorPane.IsVisible = inspector;
        NavigationSplitter.IsVisible = navigation && !pagesGrid;
        InspectorSplitter.IsVisible = inspector;
        NavigationToggle.IsChecked = navigation;
        InspectorToggle.IsChecked = inspector;
        PageViewport.ColumnDefinitions[0].MinWidth = navigation ? 200D : 0D;
        PageViewport.ColumnDefinitions[0].MaxWidth = navigation ? 320D : 0D;
        PageViewport.ColumnDefinitions[2].MinWidth = inspector ? 280D : 0D;
        PageViewport.ColumnDefinitions[2].MaxWidth = inspector ? 380D : 0D;
        PageViewport.ColumnDefinitions[0].Width = new GridLength(navigation ? _navigationWidth : 0D);
        PageViewport.ColumnDefinitions[2].Width = new GridLength(inspector ? _inspectorWidth : 0D);
    }

    private void OnNavigationToggleClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        CapturePaneWidths();
        SetPanes(NavigationToggle.IsChecked == true,
            _compactLayout == true && NavigationToggle.IsChecked == true ? false : InspectorPane.IsVisible);
        SavePaneVisibility();
    }

    private void OnInspectorToggleClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        CapturePaneWidths();
        SetPanes(_compactLayout == true && InspectorToggle.IsChecked == true ? false : NavigationPane.IsVisible,
            InspectorToggle.IsChecked == true);
        SavePaneVisibility();
    }

    private void SavePaneVisibility() => _document?.UpdatePanePreferences(_navigationWidth, _inspectorWidth,
        new StudioPanePreference(NavigationPane.IsVisible, InspectorPane.IsVisible));

    internal ListBox PagesListControl => PagesList;

    internal ListBox OrganizerListControl => OrganizerList;

    internal ListBox GridPagesListControl => GridPagesList;

    internal Button FitWidthButtonControl => FitWidthButton;

    internal Button FitPageButtonControl => FitPageButton;

    internal void FocusSearch() {
        if (_document is not null) _document.IsFocusReading = false;
        SetPanes(true, _compactLayout == true ? false : InspectorPane.IsVisible);
        SavePaneVisibility();
        NavigationTabs.SelectedIndex = 2;
        var document = _document;
        Avalonia.Threading.Dispatcher.UIThread.Post(() => {
            if (!ReferenceEquals(document, _document) || NavigationTabs.SelectedIndex != 2 || !IsEffectivelyVisible) return;
            SearchBox.Focus();
            SearchBox.SelectAll();
        }, Avalonia.Threading.DispatcherPriority.Loaded);
    }

    private async void OnSearchKeyDown(object? sender, KeyEventArgs e) {
        if (_document is null) return;
        if (e.Key == Key.Escape) {
            _document.ClearSearchCommand.Execute(null);
            (GridPagesList.IsEffectivelyVisible ? GridPagesList : PagesList).Focus();
            e.Handled = true;
        } else if (e.Key == Key.Enter) {
            e.Handled = true;
            if (_document.HasSearchResults) {
                if (e.KeyModifiers.HasFlag(KeyModifiers.Shift)) _document.PreviousSearchResultCommand.Execute(null);
                else _document.NextSearchResultCommand.Execute(null);
            } else if (_document.SearchCommand.CanExecute(null)) {
                await _document.SearchCommand.ExecuteAsync(null);
            }
        }
    }

    private void OnFillSignFlyoutOpening(object? sender, EventArgs e) => _document?.EnsureSignaturesLoaded();

    // Choosing a signature or creating one closes the menu so the next click lands on the page. The menu closes
    // after the item's command has run; closing it first would detach the item from its data context.
    private void OnFillSignItemClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
        Avalonia.Threading.Dispatcher.UIThread.Post(() => FillSignButton.Flyout?.Hide(), Avalonia.Threading.DispatcherPriority.Background);

    // Enter renames the selected bookmark; Escape restores its current title.
    private async void OnBookmarkTitleKeyDown(object? sender, KeyEventArgs e) {
        if (_document is null) return;
        if (e.Key == Key.Escape) {
            _document.BookmarkTitleDraft = _document.SelectedBookmark?.Title ?? string.Empty;
            e.Handled = true;
        } else if (e.Key == Key.Enter) {
            e.Handled = true;
            if (_document.RenameBookmarkCommand.CanExecute(null)) await _document.RenameBookmarkCommand.ExecuteAsync(null);
        }
    }

    private void OnGridPagePointerPressed(object? sender, PointerPressedEventArgs e) {
        if (!e.GetCurrentPoint(GridPagesList).Properties.IsLeftButtonPressed ||
            DataContext is not MainWindowViewModel viewModel) return;

        Control? control = e.Source as Control;
        while (control is not null) {
            if (control.DataContext is PdfPageViewModel page) {
                viewModel.SelectedPage = page;
                return;
            }
            control = control.Parent as Control;
        }
    }
}
