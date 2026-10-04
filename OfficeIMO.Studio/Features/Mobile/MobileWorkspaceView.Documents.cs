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

    private void UpdateDocumentStatus() {
        WorkspaceTitle.Text = Document?.WorkspaceMode switch {
            StudioWorkspaceMode.PdfWorkspace => Document.DocumentName,
            StudioWorkspaceMode.Home => "Home",
            StudioWorkspaceMode.Tools => "Tools",
            StudioWorkspaceMode.Convert => "Convert",
            StudioWorkspaceMode.Output => "Document output",
            StudioWorkspaceMode.DocumentHealth => "Document health",
            StudioWorkspaceMode.Provenance => "Document provenance",
            StudioWorkspaceMode.Ocr => "Recognize text",
            StudioWorkspaceMode.Invoices => "Invoices",
            StudioWorkspaceMode.Publishing => "Books",
            StudioWorkspaceMode.Jobs => "Jobs",
            StudioWorkspaceMode.Settings => "Settings",
            _ => "OfficeIMO Studio"
        };
        DocumentSubtitle.IsVisible = Document?.ShowPdfDocumentControls == true;
        DocumentSubtitle.Text = Document switch {
        { IsOpening: true } => "Opening…",
        { IsWorkspaceBusy: true } => "Working…",
        { HasDocument: true, IsDirty: true } => _controller?.IsWorkingCopy != false ? "Unsaved changes · Working copy" : "Unsaved changes",
        { HasDocument: true } => _controller?.IsWorkingCopy != false ? "Saved on this device" : "Saved",
        _ => "OfficeIMO Studio"
        };
    }

    private Task<UnsavedChangesDecision> ConfirmUnsavedChangesAsync(MainWindowViewModel document) {
        // A close request must never replace another modal decision or an in-progress note.
        if (CommandPalette.IsOpen || SheetScrim.IsVisible || DialogScrim.IsVisible || _closeDecision is not null || !this.IsAttachedToVisualTree())
            return Task.FromResult(UnsavedChangesDecision.Cancel);
        _closeDecision = new(TaskCreationOptions.RunContinuationsAsynchronously);
        CloseDescription.Text = StudioLocalization.Current.Format("Dialog.SaveChangesTo", document.DocumentName.TrimEnd(' ', '*'));
        ShowSheet(StudioLocalization.Current.Get("Dialog.UnsavedChanges"), CloseScroll,
            ShortDocumentsButton.IsVisible ? ShortDocumentsButton : DocumentsButton);
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
            source.GetSelfAndVisualAncestors().OfType<Button>().Any() ||
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

    private static void ExecuteIfAvailable(ICommand? command) {
        if (command?.CanExecute(null) == true) command.Execute(null);
    }
}
