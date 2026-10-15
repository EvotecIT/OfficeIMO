using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow {
    /// <summary>
    /// Renders the shared menu definition as the macOS menu bar. Each menu rebuilds its items from the definition when
    /// AppKit asks for an update, so recent documents, open tabs and availability are current whenever it opens.
    /// </summary>
    private void RefreshNativeMenus() {
        if (!OperatingSystem.IsMacOS()) return;
        // AppKit's exporter keeps the attached root menu identity across document switches.
        var menus = NativeMenu.GetMenu(this);
        bool attach = menus is null;
        menus ??= new NativeMenu();
        menus.Items.Clear();
        foreach (StudioMenuGroup group in BuildMenuModel(nativeMac: true)) {
            var item = (NativeMenuItem)StudioMenuRenderer.ToNative(group);
            string id = group.Id;
            item.Menu!.NeedsUpdate += (_, _) => {
                StudioMenuGroup current = BuildMenuModel(nativeMac: true).Single(candidate => candidate.Id == id);
                item.Menu.Items.Clear();
                foreach (StudioMenuEntry entry in current.Items) item.Menu.Items.Add(StudioMenuRenderer.ToNative(entry));
            };
            menus.Items.Add(item);
        }
        if (attach) NativeMenu.SetMenu(this, menus);
    }

    /// <summary>
    /// The macOS Edit commands. Native menus keep the text field as their editing target, so a focused field gets its own
    /// undo and clipboard actions; document undo must never consume text undo.
    /// </summary>
    private IEnumerable<StudioMenuEntry> TextAwareEditEntries() {
        string Text(string id) => _services.Localizer.Get("Apple." + id);
        TextBox? FocusedText() => FocusManager?.GetFocusedElement() as TextBox;
        StudioMenuCommand EditItem(string label, Key key, Action<TextBox> textAction, Func<TextBox, bool> textCanExecute,
            string? documentCommand = null, KeyModifiers extra = KeyModifiers.None) {
            var command = new RelayCommand(() => {
                if (FocusedText() is { } text) textAction(text);
                else if (documentCommand is not null) ViewModel.Commands[documentCommand].Execute(null);
            }, () => FocusedText() is { } text ? textCanExecute(text)
                : documentCommand is not null && ViewModel.Commands[documentCommand].IsAvailable);
            return new StudioMenuCommand(label, documentCommand is null ? Text(label) : ViewModel.Commands[documentCommand].Title,
                command, StudioShortcuts.Primary(key, extra));
        }
        var copy = new AsyncRelayCommand(async () => {
            if (FocusedText() is { } text) text.Copy();
            else if (FocusManager?.GetFocusedElement() is PdfPageCanvas page) await page.CopySelectedTextAsync();
        }, () => FocusedText()?.CanCopy == true || FocusManager?.GetFocusedElement() is PdfPageCanvas { HasTextSelection: true });
        var selectAll = new RelayCommand(() => {
            if (FocusedText() is { } text) text.SelectAll();
            else if (FocusManager?.GetFocusedElement() is PdfPageCanvas page) page.SelectAllPageText();
        }, () => FocusedText() is { } text ? !string.IsNullOrEmpty(text.Text) : FocusManager?.GetFocusedElement() is PdfPageCanvas);
        return [
            EditItem("Undo", Key.Z, box => box.Undo(), box => box.CanUndo, "Undo"),
            EditItem("Redo", Key.Z, box => box.Redo(), box => box.CanRedo, "Redo", KeyModifiers.Shift),
            StudioMenuSeparator.Instance,
            EditItem("Cut", Key.X, box => box.Cut(), box => box.CanCut),
            new StudioMenuCommand("Copy", Text("Copy"), copy, StudioShortcuts.Primary(Key.C)),
            EditItem("Paste", Key.V, box => box.Paste(), box => box.CanPaste),
            new StudioMenuCommand("SelectAll", Text("SelectAll"), selectAll, StudioShortcuts.Primary(Key.A))
        ];
    }
}
