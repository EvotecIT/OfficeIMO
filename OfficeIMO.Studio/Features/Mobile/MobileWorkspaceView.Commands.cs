using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Features.Mobile;

public sealed partial class MobileWorkspaceView {
    private bool _showingCommands;

    private async void OnCommandsClick(object? sender, RoutedEventArgs e) => await ShowCommandPaletteAsync();

    private async Task ShowCommandPaletteAsync() {
        if (_showingCommands || _initializing || SheetScrim.IsVisible || DialogScrim.IsVisible ||
            Document is not { } document || !this.IsAttachedToVisualTree()) return;
        _showingCommands = true;
        var previousFocus = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement() as Control;
        if (ApplicationNavigation.DisplayMode == SplitViewDisplayMode.Overlay) ApplicationNavigation.IsPaneOpen = false;
        ApplicationNavigation.IsEnabled = false;
        try {
            var command = await CommandPalette.ShowAsync(document.Commands);
            ApplicationNavigation.IsEnabled = !_initializing;
            RestoreWorkspaceFocus(previousFocus);
            if (command is not null && ReferenceEquals(Document, document)) await command.ExecuteAsync();
        } finally {
            ApplicationNavigation.IsEnabled = !_initializing && !DialogScrim.IsVisible;
            _showingCommands = false;
        }
    }

    private async void OnWorkspaceKeyDown(object? sender, KeyEventArgs e) {
        if (CommandPalette.IsOpen) return;
        if (DialogScrim.IsVisible) {
            if (e.Key == Key.Escape) { _dismissDialog?.Invoke(null); e.Handled = true; }
            return;
        }
        if (_initializing) return;
        if (e.Key == Key.Escape && SheetScrim.IsVisible) { DismissSheet(); e.Handled = true; return; }
        if (SheetScrim.IsVisible) return;
        if (e.Key == Key.Escape) {
            if (SearchPanel.IsVisible) CloseSearch();
            else if (Document?.IsFocusReading == true) Document.IsFocusReading = false;
            else if (Document?.IsAssistantVisible == true) Document.IsAssistantVisible = false;
            else if (ApplicationNavigation.DisplayMode == SplitViewDisplayMode.Overlay && ApplicationNavigation.IsPaneOpen) {
                ApplicationNavigation.IsPaneOpen = false;
                ApplicationMenuButton.Focus();
            } else return;
            e.Handled = true;
            return;
        }
        bool shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        if (e.Key == Key.Tab && e.KeyModifiers is KeyModifiers.Control or (KeyModifiers.Control | KeyModifiers.Shift)) {
            _controller?.Tabs.SelectRelativeTab(shift);
            e.Handled = true;
            return;
        }
        KeyModifiers primary = OperatingSystem.IsMacOS() || OperatingSystem.IsIOS() ? KeyModifiers.Meta : KeyModifiers.Control;
        bool primaryModifier = e.KeyModifiers == primary || e.KeyModifiers == (primary | KeyModifiers.Shift);
        if (primaryModifier) {
            switch (e.Key) {
                case Key.K when !shift:
                case Key.P when shift:
                    e.Handled = true;
                    await ShowCommandPaletteAsync();
                    return;
                case Key.O when !shift: ExecuteIfAvailable(Document?.Commands["Open"]); break;
                case Key.S: ExecuteIfAvailable(Document?.Commands[shift ? "SaveAs" : "Save"]); break;
                case Key.P when !shift: ExecuteIfAvailable(Document?.Commands["Print"]); break;
                case Key.OemComma when !shift: ExecuteIfAvailable(Document?.Commands["Settings"]); break;
                case Key.F when !shift:
                    if (IsTouchReader && Document?.HasDocument == true) OpenSearch();
                    else if (Document?.HasDocument == true) {
                        Document.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
                        if (IsTouchReader) OpenSearch(); else _editingWorkspace?.FocusSearch();
                    }
                    break;
                case Key.W when !shift && _controller is not null: await _controller.Tabs.CloseSelectedTabAsync(); break;
                case Key.T when shift && _controller is not null: await _controller.Tabs.ReopenClosedTabAsync(); break;
                case Key.OemOpenBrackets when shift: _controller?.Tabs.SelectRelativeTab(true); break;
                case Key.OemCloseBrackets when shift: _controller?.Tabs.SelectRelativeTab(false); break;
                default:
                    if (IsMobileTextEntryFocused()) return;
                    string? id = e.Key switch {
                        Key.Z => shift ? "Redo" : "Undo",
                        Key.Y when primary == KeyModifiers.Control && !shift => "Redo",
                        Key.D0 or Key.NumPad0 when !shift => "FitPage",
                        Key.D1 or Key.NumPad1 when !shift => "ActualSize",
                        Key.D2 or Key.NumPad2 when !shift => "FitWidth",
                        Key.OemPlus or Key.Add => "ZoomIn",
                        Key.OemMinus or Key.Subtract when !shift => "ZoomOut",
                        _ => null
                    };
                    if (id is null) return;
                    ExecuteIfAvailable(Document?.Commands[id]);
                    break;
            }
            e.Handled = true;
        } else if (e.Key == Key.F9 && e.KeyModifiers == KeyModifiers.None && !IsMobileTextEntryFocused()) {
            ExecuteIfAvailable(Document?.Commands["FocusReading"]);
            e.Handled = true;
        }
    }

    private bool IsMobileTextEntryFocused() =>
        TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement() is Visual focused &&
        focused.GetSelfAndVisualAncestors().Any(control => control is TextBox or ComboBox or NumericUpDown);
}
