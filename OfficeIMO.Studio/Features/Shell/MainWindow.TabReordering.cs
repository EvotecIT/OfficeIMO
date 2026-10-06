using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow {
    private StudioDocumentTabViewModel? _draggedDocumentTab;
    private Point _tabPressPosition;
    private IPointer? _tabDragPointer;
    private IPointer? _tabPressPointer;
    private TabItem? _tabDropTarget;
    private bool _tabDropAfter;
    private bool _acquiringTabCapture;

    private void InitializeDocumentTabInteractions() {
        ((MenuFlyout)DocumentListButton.Flyout!).ItemsSource = _documentMenuItems;
        DocumentTabs.AddHandler(ContextRequestedEvent, OnDocumentTabContextRequested, RoutingStrategies.Tunnel);
        DocumentTabs.AddHandler(KeyDownEvent, OnDocumentTabsKeyDown, RoutingStrategies.Tunnel);
        DocumentTabs.AddHandler(PointerPressedEvent, OnTabDragPressed, RoutingStrategies.Tunnel);
        DocumentTabs.AddHandler(PointerMovedEvent, OnTabDragMoved, handledEventsToo: true);
        DocumentTabs.AddHandler(PointerReleasedEvent, OnTabDragReleased, handledEventsToo: true);
        DocumentTabs.PointerCaptureLost += (_, e) => {
            if (!_acquiringTabCapture && e.Pointer == _tabPressPointer) ClearDocumentTabDrag();
        };
    }

    private void OnTabDragPressed(object? sender, PointerPressedEventArgs e) {
        ClearDocumentTabDrag();
        if (e.Source is not Visual source ||
            source.GetSelfAndVisualAncestors().OfType<Button>().Any()) return;
        var container = source.GetSelfAndVisualAncestors().OfType<TabItem>().FirstOrDefault();
        if (container?.DataContext is not StudioDocumentTabViewModel tab) return;
        if (!e.GetCurrentPoint(DocumentTabs).Properties.IsLeftButtonPressed) return;
        _draggedDocumentTab = tab;
        _tabPressPointer = e.Pointer;
        _tabPressPosition = e.GetPosition(DocumentTabs);
    }

    private void OnDocumentTabContextRequested(object? sender, ContextRequestedEventArgs e) {
        if (e.Source is not Visual source) return;
        var container = source.GetSelfAndVisualAncestors().OfType<TabItem>().FirstOrDefault();
        if (container?.DataContext is not StudioDocumentTabViewModel tab) return;
        ClearDocumentTabDrag();
        var menu = new MenuFlyout();
        AddDocumentTabActions(menu.Items, tab);
        container.ContextFlyout = menu;
        menu.ShowAt(container, e.TryGetPosition(container, out _));
        e.Handled = true;
    }

    private void OnTabDragMoved(object? sender, PointerEventArgs e) {
        if (_draggedDocumentTab is null || e.Pointer != _tabPressPointer) return;
        Point position = e.GetPosition(DocumentTabs);
        // Release and capture loss end the gesture; native move events may omit button flags.
        if (_tabDragPointer is null) {
            if (Math.Abs(position.X - _tabPressPosition.X) < 8) return;
            _tabDragPointer = e.Pointer;
            _acquiringTabCapture = true;
            try { e.Pointer.Capture(DocumentTabs); }
            finally { _acquiringTabCapture = false; }
            DocumentTabs.Classes.Add("reordering");
        }
        UpdateTabDropTarget(position);
        e.Handled = true;
    }

    private void UpdateTabDropTarget(Point position) {
        _tabDropTarget?.Classes.Remove("dropBefore");
        _tabDropTarget?.Classes.Remove("dropAfter");
        _tabDropTarget = null;
        if (!new Rect(DocumentTabs.Bounds.Size).Contains(position)) return;
        for (int index = 0; index < TabHost.Tabs.Count; index++) {
            if (DocumentTabs.ContainerFromIndex(index) is not TabItem item ||
                item.TranslatePoint(default, DocumentTabs) is not { } origin) continue;
            if (position.X < origin.X || position.X > origin.X + item.Bounds.Width) continue;
            _tabDropTarget = item;
            _tabDropAfter = position.X >= origin.X + item.Bounds.Width / 2;
            item.Classes.Add(_tabDropAfter ? "dropAfter" : "dropBefore");
            break;
        }
    }

    private void OnTabDragReleased(object? sender, PointerReleasedEventArgs e) {
        if (e.InitialPressMouseButton != MouseButton.Left || e.Pointer != _tabPressPointer) return;
        var dragged = _draggedDocumentTab;
        if (_tabDragPointer is not null && dragged is not null) {
            UpdateTabDropTarget(e.GetPosition(DocumentTabs));
            if (_tabDropTarget?.DataContext is StudioDocumentTabViewModel target) {
                int oldIndex = TabHost.Tabs.IndexOf(dragged);
                int insertion = TabHost.Tabs.IndexOf(target) + (_tabDropAfter ? 1 : 0);
                if (oldIndex < insertion) insertion--;
                TabHost.MoveTab(dragged, insertion);
                FocusSelectedDocumentTab();
            }
            e.Handled = true;
        }
        ClearDocumentTabDrag();
    }

    private void ClearDocumentTabDrag() {
        _tabDropTarget?.Classes.Remove("dropBefore");
        _tabDropTarget?.Classes.Remove("dropAfter");
        _tabDropTarget = null;
        _draggedDocumentTab = null;
        _tabPressPointer = null;
        DocumentTabs.Classes.Remove("reordering");
        var pointer = _tabDragPointer;
        _tabDragPointer = null;
        if (pointer?.Captured == DocumentTabs) pointer.Capture(null);
    }
}
