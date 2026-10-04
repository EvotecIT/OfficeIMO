using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using OfficeIMO.Studio.Features.Organizer;

namespace OfficeIMO.Studio.Features.Workspace;

public sealed partial class DocumentWorkspaceView {
    private static readonly DataFormat<string> OrganizerPageFormat =
        DataFormat.CreateInProcessFormat<string>("officeimo-studio-organizer-page");
    private PointerPressedEventArgs? _organizerDragPress;
    private PdfOrganizerPageViewModel? _organizerDragPage;
    private Point _organizerDragStart;
    private bool _organizerDragStarted;

    private void InitializeOrganizerInput() {
        OrganizerList.AddHandler(KeyDownEvent, OnOrganizerKeyDown, Avalonia.Interactivity.RoutingStrategies.Tunnel);
        OrganizerList.AddHandler(PointerPressedEvent, OnOrganizerPointerPressed, handledEventsToo: true);
        OrganizerList.AddHandler(PointerMovedEvent, OnOrganizerPointerMoved, handledEventsToo: true);
        OrganizerList.AddHandler(PointerReleasedEvent, OnOrganizerPointerReleased, handledEventsToo: true);
        OrganizerList.AddHandler(DragDrop.DragOverEvent, OnOrganizerDragOver);
        OrganizerList.AddHandler(DragDrop.DropEvent, OnOrganizerDrop);
        DataContextChanged += (_, _) => ClearOrganizerDrag();
        DetachedFromVisualTree += (_, _) => ClearOrganizerDrag();
    }

    private void OnOrganizerPointerPressed(object? sender, PointerPressedEventArgs e) {
        if (_document is null || !e.GetCurrentPoint(OrganizerList).Properties.IsLeftButtonPressed) return;
        _organizerDragPage = FindOrganizerPage(e.Source);
        if (_organizerDragPage is null) return;
        _document!.NavigateToOrganizerPage(_organizerDragPage.PageNumber);
        if (e.Pointer.Type != PointerType.Mouse) return;
        _organizerDragPress = e;
        _organizerDragStart = e.GetPosition(OrganizerList);
        _organizerDragStarted = false;
    }

    private async void OnOrganizerPointerMoved(object? sender, PointerEventArgs e) {
        if (_organizerDragStarted || _organizerDragPress is null || _organizerDragPage is null ||
            !e.GetCurrentPoint(OrganizerList).Properties.IsLeftButtonPressed) return;
        Point current = e.GetPosition(OrganizerList);
        if (Math.Abs(current.X - _organizerDragStart.X) < 6D && Math.Abs(current.Y - _organizerDragStart.Y) < 6D) return;

        _organizerDragStarted = true;
        var transfer = new DataTransfer();
        transfer.Add(DataTransferItem.Create(
            OrganizerPageFormat,
            _organizerDragPage.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        PointerPressedEventArgs press = _organizerDragPress;
        ClearOrganizerDrag();
        await DragDrop.DoDragDropAsync(press, transfer, DragDropEffects.Move);
    }

    private void OnOrganizerPointerReleased(object? sender, PointerReleasedEventArgs e) => ClearOrganizerDrag();

    private async void OnOrganizerKeyDown(object? sender, KeyEventArgs e) {
        if (_document is null) return;
        if (_document!.IsPagesDocumentMode && e.KeyModifiers == KeyModifiers.Alt && e.Key is Key.Up or Key.Down) {
            e.Handled = true;
            if (!_document!.CanMutateSelection) return;
            await (e.Key == Key.Up ? _document!.MoveSelectedUpCommand : _document!.MoveSelectedDownCommand).ExecuteAsync(null);
            if (_document!.OrganizerPages.FirstOrDefault(page => page.IsSelected) is { } selected) OrganizerList.ScrollIntoView(selected);
            OrganizerList.Focus();
            return;
        }
        if (e.Key is not (Key.Enter or Key.Space)) return;
        PdfOrganizerPageViewModel? page = FindOrganizerPage(e.Source)
            ?? OrganizerList.SelectedItem as PdfOrganizerPageViewModel;
        if (page is null) return;
        _document!.NavigateToOrganizerPage(page.PageNumber);
    }

    private void OnOrganizerDragOver(object? sender, DragEventArgs e) {
        e.DragEffects = TryGetOrganizerPage(e, out _) && FindOrganizerPage(e.Source) is not null
            ? DragDropEffects.Move
            : DragDropEffects.None;
        e.Handled = true;
    }

    private async void OnOrganizerDrop(object? sender, DragEventArgs e) {
        e.Handled = true;
        PdfOrganizerPageViewModel? target = FindOrganizerPage(e.Source);
        if (_document is not null && target is not null && TryGetOrganizerPage(e, out int draggedPage)) {
            await _document!.ReorderByDropAsync(draggedPage, target.PageNumber);
        }
    }

    private static bool TryGetOrganizerPage(DragEventArgs e, out int pageNumber) {
        foreach (IDataTransferItem item in e.DataTransfer.Items) {
            string? value = item.TryGetValue(OrganizerPageFormat);
            if (int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out pageNumber)) {
                return true;
            }
        }
        pageNumber = 0;
        return false;
    }

    private static PdfOrganizerPageViewModel? FindOrganizerPage(object? source) {
        Control? control = source as Control;
        while (control is not null) {
            if (control.DataContext is PdfOrganizerPageViewModel page) return page;
            control = control.Parent as Control;
        }
        return null;
    }

    private void ClearOrganizerDrag() {
        _organizerDragPress = null;
        _organizerDragPage = null;
        _organizerDragStarted = false;
    }
}
