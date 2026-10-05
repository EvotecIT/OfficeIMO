using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow {
    // The flyout presenter keeps this collection while each opening refreshes its entries.
    private readonly System.Collections.ObjectModel.ObservableCollection<object> _documentMenuItems = new();

    private Task CloseSelectedDocumentTabAsync() => CloseDocumentTabAsync(TabHost.SelectedTab);

    private async void OnDocumentTabCloseClick(object? sender, RoutedEventArgs e) {
        if (sender is Control { DataContext: StudioDocumentTabViewModel tab })
            await CloseDocumentTabAsync(tab);
    }

    private async Task CloseDocumentTabAsync(StudioDocumentTabViewModel? tab) {
        if (tab is null || !tab.CloseCommand.CanExecute(null)) return;
        bool restoreTabFocus = DocumentTabs.IsKeyboardFocusWithin || CompactCloseDocumentButton.IsKeyboardFocusWithin ||
                               CompactDocumentPicker.IsKeyboardFocusWithin;
        await tab.CloseCommand.ExecuteAsync(null);
        if (!restoreTabFocus || _windowClosed || TabHost.Tabs.Contains(tab)) return;
        // A removed close button cannot retain keyboard focus. Keep subsequent tab actions usable.
        FocusSelectedDocumentTab();
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
        if (e.Key == Key.Escape && _draggedDocumentTab is not null) {
            ClearDocumentTabDrag();
            e.Handled = true;
            return;
        }
        if (e.KeyModifiers == (KeyModifiers.Alt | KeyModifiers.Shift) && e.Key is Key.Left or Key.Right) {
            if (TabHost.SelectedTab is { } tab) {
                TabHost.MoveTab(tab, TabHost.Tabs.IndexOf(tab) + (e.Key == Key.Left ? -1 : 1));
                FocusSelectedDocumentTab();
            }
            e.Handled = true;
            return;
        }
        if (e.KeyModifiers != KeyModifiers.None || e.Key is not (Key.Home or Key.End) || !TabHost.HasTabs) return;
        DocumentTabs.SelectedIndex = e.Key == Key.Home ? 0 : TabHost.Tabs.Count - 1;
        DocumentTabs.ContainerFromIndex(DocumentTabs.SelectedIndex)?.Focus(NavigationMethod.Directional);
        e.Handled = true;
    }

    private void FocusSelectedDocumentTab() => Dispatcher.UIThread.Post(() => {
        if (_windowClosed) return;
        Control next = CompactDocumentPicker.IsEffectivelyVisible ? CompactDocumentPicker :
            DocumentTabs.ContainerFromIndex(DocumentTabs.SelectedIndex) ?? OpenDocumentTabButton;
        next.Focus(NavigationMethod.Directional);
    }, DispatcherPriority.Loaded);

    private void OnDocumentListOpening(object? sender, EventArgs e) {
        if (sender is not MenuFlyout menu) return;
        _documentMenuItems.Clear();
        foreach (var tab in TabHost.Tabs) {
            var item = new MenuItem {
                Header = tab.Title,
                ToggleType = MenuItemToggleType.Radio,
                IsChecked = ReferenceEquals(tab, TabHost.SelectedTab)
            };
            item.Click += (_, _) => {
                if (TabHost.Tabs.Contains(tab)) {
                    TabHost.SelectedTab = tab;
                    FocusSelectedDocumentTab();
                }
            };
            _documentMenuItems.Add(item);
        }
        _documentMenuItems.Add(new Separator());
        if (TabHost.SelectedTab is { } selected) AddDocumentTabActions(_documentMenuItems, selected);
        var reopen = new MenuItem {
            Header = Infrastructure.Localization.StudioLocalization.Current.Get("Apple.ReopenTab"),
            IsEnabled = TabHost.CanReopenClosedTab
        };
        reopen.Click += async (_, _) => await TabHost.ReopenClosedTabAsync();
        _documentMenuItems.Add(reopen);

    }

    private void AddDocumentTabActions(System.Collections.IList items, StudioDocumentTabViewModel tab) {
        var strings = Infrastructure.Localization.StudioLocalization.Current;
        foreach (int offset in new[] { -1, 1 }) {
            var move = new MenuItem {
                Header = strings.Get(offset < 0 ? "Tabs.MoveLeft" : "Tabs.MoveRight"),
                IsEnabled = TabHost.Tabs.IndexOf(tab) + offset >= 0 && TabHost.Tabs.IndexOf(tab) + offset < TabHost.Tabs.Count
            };
            move.Click += (_, _) => {
                TabHost.MoveTab(tab, TabHost.Tabs.IndexOf(tab) + offset);
                FocusSelectedDocumentTab();
            };
            items.Add(move);
        }
        var close = new MenuItem { Header = tab.CloseLabel };
        close.Click += async (_, _) => {
            await CloseDocumentTabAsync(tab);
            if (!TabHost.Tabs.Contains(tab)) FocusSelectedDocumentTab();
        };
        items.Add(close);
    }
}
