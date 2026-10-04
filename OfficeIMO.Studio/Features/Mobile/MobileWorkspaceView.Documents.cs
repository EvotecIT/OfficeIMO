using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Localization;
using System.Windows.Input;

namespace OfficeIMO.Studio.Features.Mobile;

public sealed partial class MobileWorkspaceView {
    private TaskCompletionSource<UnsavedChangesDecision>? _closeDecision;

    private void UpdateDocumentStatus() => DocumentSubtitle.Text = Document switch {
        { IsOpening: true } => "Opening…",
        { IsWorkspaceBusy: true } => "Working…",
        { HasDocument: true, IsDirty: true } => "Unsaved changes · Working copy",
        { HasDocument: true } => "Saved on this device",
        _ => "OfficeIMO Studio"
    };

    private Task<UnsavedChangesDecision> ConfirmUnsavedChangesAsync(MainWindowViewModel document) {
        // A close request must never replace another modal decision or an in-progress note.
        if (SheetScrim.IsVisible || _closeDecision is not null || !this.IsAttachedToVisualTree())
            return Task.FromResult(UnsavedChangesDecision.Cancel);
        _closeDecision = new(TaskCreationOptions.RunContinuationsAsynchronously);
        CloseDescription.Text = StudioLocalization.Current.Format("Dialog.SaveChangesTo", document.DocumentName.TrimEnd(' ', '*'));
        ShowSheet(StudioLocalization.Current.Get("Dialog.UnsavedChanges"), CloseScroll, DocumentsButton);
        return _closeDecision.Task;
    }

    private void CompleteCloseDecision(UnsavedChangesDecision decision) {
        if (_closeDecision is not { } pending) return;
        _closeDecision = null;
        DismissSheet();
        pending.TrySetResult(decision);
    }

    private void OnSaveAndCloseClick(object? sender, RoutedEventArgs e) => CompleteCloseDecision(UnsavedChangesDecision.Save);
    private void OnDiscardAndCloseClick(object? sender, RoutedEventArgs e) => CompleteCloseDecision(UnsavedChangesDecision.Discard);

    private void OnDocumentsClick(object? sender, RoutedEventArgs e) {
        if (_controller is null) return;
        ShowSheet("Open documents", DocumentList, sender as Control);
        DocumentList.ScrollIntoView(_controller.Tabs.SelectedTab!);
    }

    private void OnDocumentTapped(object? sender, TappedEventArgs e) {
        if (e.Source is not Avalonia.Visual source ||
            source.GetSelfAndVisualAncestors().OfType<ListBoxItem>().FirstOrDefault()?.DataContext is not StudioDocumentTabViewModel tab ||
            _controller is null) return;
        _controller.Tabs.SelectedTab = tab;
        DismissSheet();
        RevealSelectedTab(focus: true);
        e.Handled = true;
    }

    private void OnDocumentListKeyDown(object? sender, KeyEventArgs e) {
        if (e.Key != Key.Enter) return;
        DismissSheet();
        RevealSelectedTab(focus: true);
        e.Handled = true;
    }

    private void OnWorkspaceKeyDown(object? sender, KeyEventArgs e) {
        if (e.Key == Key.Escape && SheetScrim.IsVisible) { DismissSheet(); e.Handled = true; return; }
        if (SheetScrim.IsVisible) return;
        if (e.Key == Key.Escape && SearchPanel.IsVisible) { CloseSearch(); e.Handled = true; return; }
        bool primary = e.KeyModifiers.HasFlag(KeyModifiers.Meta) || e.KeyModifiers.HasFlag(KeyModifiers.Control);
        bool shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        if (!primary || e.KeyModifiers.HasFlag(KeyModifiers.Alt)) return;
        switch (e.Key) {
            case Key.O when !shift: ExecuteIfAvailable(Document?.OpenCommand); break;
            case Key.S when !shift: ExecuteIfAvailable(Document?.SaveCommand); break;
            case Key.F when !shift: OpenSearch(); break;
            case Key.W when !shift && _controller is not null: _ = _controller.Tabs.CloseSelectedTabAsync(); break;
            case Key.T when shift && _controller is not null: _ = _controller.Tabs.ReopenClosedTabAsync(); break;
            case Key.Tab: _controller?.Tabs.SelectRelativeTab(shift); break;
            case Key.OemOpenBrackets when shift: _controller?.Tabs.SelectRelativeTab(true); break;
            case Key.OemCloseBrackets when shift: _controller?.Tabs.SelectRelativeTab(false); break;
            case Key.Z when !MobileSearchBox.IsKeyboardFocusWithin:
                ExecuteIfAvailable(shift ? Document?.RedoCommand : Document?.UndoCommand);
                break;
            default: return;
        }
        e.Handled = true;
    }

    private static void ExecuteIfAvailable(ICommand? command) {
        if (command?.CanExecute(null) == true) command.Execute(null);
    }
}
