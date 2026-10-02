using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow {
    /// <summary>Rebinds native menus when the selected document changes; command guards remain in the catalog.</summary>
    private void RefreshNativeMenus() {
        if (!OperatingSystem.IsMacOS()) return;
        string Text(string id) => _services.Localizer.Get("Apple." + id);
        NativeMenuItem Item(string id, string? gesture = null) => new() {
            Header = ViewModel.Commands[id].Title,
            Command = ViewModel.Commands[id],
            Gesture = gesture is null ? null : KeyGesture.Parse(gesture)
        };
        NativeMenuItem Action(string title, Action action, string? gesture = null) => new() {
            Header = Text(title), Command = new RelayCommand(action),
            Gesture = gesture is null ? null : KeyGesture.Parse(gesture)
        };
        NativeMenuItem Group(string name, params NativeMenuItemBase[] items) {
            var menu = new NativeMenu();
            foreach (var item in items) menu.Items.Add(item);
            return new NativeMenuItem(Text(name)) { Menu = menu };
        }

        var close = new AsyncRelayCommand(TabHost.CloseSelectedTabAsync, () => TabHost.HasTabs);
        var reopen = new AsyncRelayCommand(TabHost.ReopenClosedTabAsync);
        var open = Item("Open", "Meta+O"); open.Header = Text("Open");
        var saveAs = Item("SaveAs", "Meta+Shift+S"); saveAs.Header = Text("SaveAs");
        var file = Group("File", open, new NativeMenuItemSeparator(), Item("Save", "Meta+S"), saveAs,
            new NativeMenuItemSeparator(), Item("Print", "Meta+P"), new NativeMenuItemSeparator(),
            new NativeMenuItem(Text("CloseDocument")) { Command = close, Gesture = KeyGesture.Parse("Meta+W") });
        file.Menu!.NeedsUpdate += (_, _) => close.NotifyCanExecuteChanged();

        // Native menus retain the text field's editing target. Document undo must never consume text undo.
        TextBox? FocusedText() => FocusManager?.GetFocusedElement() as TextBox;
        var editing = new List<IRelayCommand>();
        NativeMenuItem EditItem(string label, string gesture, Action<TextBox> textAction,
            Func<TextBox, bool> textCanExecute, string? documentCommand = null) {
            var command = new RelayCommand(() => {
                if (FocusedText() is { } text) textAction(text);
                else if (documentCommand is not null) ViewModel.Commands[documentCommand].Execute(null);
            }, () => FocusedText() is { } text ? textCanExecute(text)
                : documentCommand is not null && ViewModel.Commands[documentCommand].IsAvailable);
            editing.Add(command);
            return new NativeMenuItem(documentCommand is null ? Text(label) : ViewModel.Commands[documentCommand].Title) {
                Command = command, Gesture = KeyGesture.Parse(gesture)
            };
        }
        var copy = new AsyncRelayCommand(async () => {
            if (FocusedText() is { } text) text.Copy();
            else if (FocusManager?.GetFocusedElement() is PdfPageCanvas page) await page.CopySelectedTextAsync();
        }, () => FocusedText()?.CanCopy == true || FocusManager?.GetFocusedElement() is PdfPageCanvas { HasTextSelection: true });
        var selectAll = new RelayCommand(() => {
            if (FocusedText() is { } text) text.SelectAll();
            else if (FocusManager?.GetFocusedElement() is PdfPageCanvas page) page.SelectAllPageText();
        }, () => FocusedText() is { } text ? !string.IsNullOrEmpty(text.Text) : FocusManager?.GetFocusedElement() is PdfPageCanvas);
        editing.Add(copy);
        editing.Add(selectAll);
        var edit = Group("Edit",
            EditItem("Undo", "Meta+Z", box => box.Undo(), box => box.CanUndo, "Undo"),
            EditItem("Redo", "Meta+Shift+Z", box => box.Redo(), box => box.CanRedo, "Redo"),
            new NativeMenuItemSeparator(),
            EditItem("Cut", "Meta+X", box => box.Cut(), box => box.CanCut),
            new NativeMenuItem(Text("Copy")) { Command = copy, Gesture = KeyGesture.Parse("Meta+C") },
            EditItem("Paste", "Meta+V", box => box.Paste(), box => box.CanPaste),
            new NativeMenuItem(Text("SelectAll")) { Command = selectAll, Gesture = KeyGesture.Parse("Meta+A") },
            new NativeMenuItemSeparator(), Action("Find", FocusDocumentSearch, "Meta+F"));
        edit.Menu!.NeedsUpdate += (_, _) => { foreach (var command in editing) command.NotifyCanExecuteChanged(); };

        var view = Group("View", Action("ToggleSidebar", () => OnSidebarToggleClick(null, new()), "Meta+Alt+S"),
            new NativeMenuItem(Text("Commands")) { Command = new AsyncRelayCommand(ShowCommandPaletteAsync), Gesture = KeyGesture.Parse("Meta+K") },
            new NativeMenuItemSeparator(), Item("Home"), Item("Tools"), Item("Jobs"),
            new NativeMenuItemSeparator(), Item("ZoomIn", "Meta+OemPlus"), Item("ZoomOut", "Meta+OemMinus"),
            Item("ActualSize", "Meta+D1"), Item("FitPage", "Meta+D0"), Item("FitWidth", "Meta+D2"), Item("FocusReading"));
        var fullscreen = Action("FullScreen", () => WindowState = WindowState == WindowState.FullScreen ? WindowState.Normal : WindowState.FullScreen, "Meta+Control+F");
        var window = Group("Window", Action("Minimize", () => WindowState = WindowState.Minimized, "Meta+M"), fullscreen,
            new NativeMenuItemSeparator(), Action("NextTab", () => TabHost.SelectRelativeTab(false), "Control+Tab"),
            Action("PreviousTab", () => TabHost.SelectRelativeTab(true), "Control+Shift+Tab"),
            new NativeMenuItem(Text("ReopenTab")) { Command = reopen, Gesture = KeyGesture.Parse("Meta+Shift+T") });
        window.Menu!.NeedsUpdate += (_, _) => fullscreen.Header = Text(WindowState == WindowState.FullScreen ? "LeaveFullScreen" : "FullScreen");
        // AppKit's exporter keeps the attached root menu identity across document switches.
        var menus = NativeMenu.GetMenu(this);
        bool attach = menus is null;
        menus ??= new NativeMenu();
        menus.Items.Clear();
        foreach (var menu in new[] { file, edit, view, window }) menus.Items.Add(menu);
        if (attach) NativeMenu.SetMenu(this, menus);
    }
}
