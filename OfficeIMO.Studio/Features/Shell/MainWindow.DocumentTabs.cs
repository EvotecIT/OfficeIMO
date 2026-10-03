using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow {
    private Task CloseSelectedDocumentTabAsync() => CloseDocumentTabAsync(TabHost.SelectedTab);

    private async void OnDocumentTabCloseClick(object? sender, RoutedEventArgs e) {
        if (sender is Control { DataContext: StudioDocumentTabViewModel tab })
            await CloseDocumentTabAsync(tab);
    }

    private async Task CloseDocumentTabAsync(StudioDocumentTabViewModel? tab) {
        if (tab is null || !tab.CloseCommand.CanExecute(null)) return;
        bool restoreTabFocus = DocumentTabs.IsKeyboardFocusWithin;
        await tab.CloseCommand.ExecuteAsync(null);
        if (!restoreTabFocus || _windowClosed || TabHost.Tabs.Contains(tab)) return;
        // A removed close button cannot retain keyboard focus. Keep subsequent tab actions usable.
        Dispatcher.UIThread.Post(() => {
            if (_windowClosed) return;
            Control next = DocumentTabs.ContainerFromIndex(DocumentTabs.SelectedIndex) ?? OpenDocumentTabButton;
            next.Focus(NavigationMethod.Tab);
        }, DispatcherPriority.Loaded);
    }

    private void OnDocumentTabSelectionChanged(object? sender, SelectionChangedEventArgs e) {
        if (!ReferenceEquals(e.Source, DocumentTabs)) return;
        RevealSelectedDocumentTab();
    }

    private void RevealSelectedDocumentTab() {
        // Selection may come from a menu or shortcut while the tab is outside the viewport.
        Dispatcher.UIThread.Post(() => {
            if (!_windowClosed && DocumentTabs.IsEffectivelyVisible)
                DocumentTabs.ContainerFromIndex(DocumentTabs.SelectedIndex)?.BringIntoView();
        }, DispatcherPriority.Loaded);
    }

    private void OnDocumentTabsKeyDown(object? sender, KeyEventArgs e) {
        if (e.KeyModifiers != KeyModifiers.None || e.Key is not (Key.Home or Key.End) || !TabHost.HasTabs) return;
        DocumentTabs.SelectedIndex = e.Key == Key.Home ? 0 : TabHost.Tabs.Count - 1;
        DocumentTabs.ContainerFromIndex(DocumentTabs.SelectedIndex)?.Focus(NavigationMethod.Directional);
        e.Handled = true;
    }

    private void OnDocumentListClick(object? sender, RoutedEventArgs e) {
        var menu = new MenuFlyout();
        foreach (var tab in TabHost.Tabs) {
            var item = new MenuItem {
                Header = tab.Title,
                ToggleType = MenuItemToggleType.Radio,
                IsChecked = ReferenceEquals(tab, TabHost.SelectedTab)
            };
            item.Click += (_, _) => {
                if (TabHost.Tabs.Contains(tab)) TabHost.SelectedTab = tab;
            };
            menu.Items.Add(item);
        }
        menu.ShowAt(DocumentListButton);
    }
}
