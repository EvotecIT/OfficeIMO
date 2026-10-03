using Avalonia;
using Avalonia.Controls;
using Avalonia.VisualTree;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Features.Mobile;

public sealed partial class MobileWorkspaceView {
    private MobileDocumentController? _controller;
    private bool _selectingThumbnail;
    // One thumbnail view per page: moving the list between hosts preserves the renderer's viewport lifetime.
    private readonly ListBox _pageList = new() {
        Classes = { "pages" },
        ItemTemplate = new FuncDataTemplate<PdfOrganizerPageViewModel>((page, _) => new PdfOrganizerPageView { DataContext = page })
    };

    internal void Connect(MobileDocumentController controller) {
        if (_controller is not null) _controller.ActiveDocumentChanged -= OnActiveDocumentChanged;
        _controller = controller;
        _controller.ActiveDocumentChanged += OnActiveDocumentChanged;
        MobileTabs.DataContext = controller.Tabs;
        TabBar.IsVisible = true;
        DataContext = controller.Document;
    }

    private void InitializeNavigation() {
        MobileTabs.AddHandler(ContextRequestedEvent, OnTabContextRequested, RoutingStrategies.Tunnel);
        ScrollViewer.SetHorizontalScrollBarVisibility(_pageList, Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled);
        _pageList.SelectionChanged += (_, _) => {
            if (_selectingThumbnail || _pageList.SelectedItem is not PdfOrganizerPageViewModel page || Document is not { } document) return;
            document.SelectedPage = document.Pages.FirstOrDefault(item => item.PageNumber == page.PageNumber);
            PageScroll.Offset = default;
            if (SheetScrim.IsVisible && SheetPages.IsVisible) DismissSheet();
        };
    }

    private void OnActiveDocumentChanged(object? sender, EventArgs e) {
        if (_controller is null) return;
        HostPageList(null);
        DataContext = _controller.Document;
        TabBar.IsVisible = _controller.Tabs.HasTabs;
        // Closed documents should not remain retained by presentation-only state.
        foreach (var document in _configuredDocuments.Where(document => !_controller.Tabs.OperationDocuments.Contains(document)).ToArray()) {
            _configuredDocuments.Remove(document);
            _noteDrafts.Remove(document);
        }
        UpdateLayoutMode();
    }

    private void RefreshPageList() {
        _selectingThumbnail = true;
        try {
            _pageList.ItemsSource = Document?.OrganizerPages;
            SelectCurrentThumbnail();
        } finally { _selectingThumbnail = false; }
        UpdateLayoutMode();
    }

    private void SelectCurrentThumbnail() {
        bool previous = _selectingThumbnail;
        _selectingThumbnail = true;
        try {
            _pageList.SelectedItem = Document?.OrganizerPages.FirstOrDefault(page => page.PageNumber == Document.SelectedPage?.PageNumber);
            if (_pageList.SelectedItem is { } selected) _pageList.ScrollIntoView(selected);
        } finally { _selectingThumbnail = previous; }
    }

    private void HostPageList(ContentControl? host) {
        if (ReferenceEquals(host?.Content, _pageList)) return;
        SidebarPages.Content = null;
        SheetPages.Content = null;
        _pageList.Height = double.NaN;
        if (host is not null) host.Content = _pageList;
    }

    private async void OnCloseTabClick(object? sender, RoutedEventArgs e) {
        if (sender is not Control { DataContext: StudioDocumentTabViewModel tab } || _controller is null) return;
        e.Handled = true;
        await tab.CloseCommand.ExecuteAsync(null);
        RevealSelectedTab(focus: true);
    }

    private void OnTabSelectionChanged(object? sender, SelectionChangedEventArgs e) {
        if (ReferenceEquals(e.Source, MobileTabs)) RevealSelectedTab();
    }

    private void RevealSelectedTab(bool focus = false) => Dispatcher.UIThread.Post(() => {
        var selected = MobileTabs.ContainerFromIndex(MobileTabs.SelectedIndex);
        selected?.BringIntoView();
        if (focus) (selected ?? OpenButton).Focus(NavigationMethod.Directional);
    }, DispatcherPriority.Loaded);

    private void OnTabContextRequested(object? sender, ContextRequestedEventArgs e) {
        if (_controller is null || e.Source is not Visual source ||
            source.GetSelfAndVisualAncestors().OfType<TabItem>().FirstOrDefault() is not { DataContext: StudioDocumentTabViewModel tab } container) return;
        var menu = new MenuFlyout();
        foreach (int offset in new[] { -1, 1 }) {
            var item = new MenuItem { Header = offset < 0 ? "Move tab left" : "Move tab right",
                IsEnabled = _controller.Tabs.Tabs.IndexOf(tab) + offset >= 0 && _controller.Tabs.Tabs.IndexOf(tab) + offset < _controller.Tabs.Tabs.Count };
            item.Click += (_, _) => { _controller.Tabs.MoveTab(tab, _controller.Tabs.Tabs.IndexOf(tab) + offset); RevealSelectedTab(); };
            menu.Items.Add(item);
        }
        menu.Items.Add(new Separator());
        menu.Items.Add(new MenuItem { Header = "Close document", Command = tab.CloseCommand });
        container.ContextFlyout = menu;
        menu.ShowAt(container);
        e.Handled = true;
    }

    private void OnTabsKeyDown(object? sender, KeyEventArgs e) {
        if (_controller is null) return;
        if (e.KeyModifiers == (KeyModifiers.Alt | KeyModifiers.Shift) && e.Key is Key.Left or Key.Right && _controller.Tabs.SelectedTab is { } tab) {
            _controller.Tabs.MoveTab(tab, _controller.Tabs.Tabs.IndexOf(tab) + (e.Key == Key.Left ? -1 : 1));
            RevealSelectedTab(focus: true);
            e.Handled = true;
        }
    }
}
