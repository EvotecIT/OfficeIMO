using Avalonia.Controls;
using Avalonia.Input;
using System.ComponentModel;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
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
        bool compact = width < 1100D;
        if (_compactLayout == compact) return;
        _compactLayout = compact;
        DocumentModeButtons.IsVisible = !compact;
        DocumentModePicker.IsVisible = compact;
        ShowContextPanes();
    }

    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName == nameof(MainWindowViewModel.DocumentViewState)) RestorePaneWidths();
        else if (e.PropertyName is nameof(MainWindowViewModel.DocumentMode) or nameof(MainWindowViewModel.IsFocusReading)) ShowContextPanes();
        else if (e.PropertyName == nameof(MainWindowViewModel.SelectedObject) && _document?.SelectedObject is not null)
            SetPanes(_compactLayout != true && NavigationPane.IsVisible, true);
    }

    private void ShowContextPanes() {
        if (_document?.IsFocusReading == true) { SetPanes(false, false); return; }
        StudioDocumentMode mode = _document?.DocumentMode ?? StudioDocumentMode.View;
        bool inspector = mode is not StudioDocumentMode.View and not StudioDocumentMode.Pages;
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
        NavigationPane.IsVisible = navigation;
        InspectorPane.IsVisible = inspector;
        NavigationSplitter.IsVisible = navigation;
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
        SearchBox.Focus();
        SearchBox.SelectAll();
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
