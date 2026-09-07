using Avalonia.Controls;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PageImportDialog : Window {
    public PageImportDialog() { InitializeComponent(); }
    internal PageImportDialog(PageImportPreviewViewModel model) : this() { DataContext = model; Opened += (_, _) => InsertionInput.Focus(); }
    private void OnCancel(object? sender, RoutedEventArgs args) => Close(false);
    private void OnApply(object? sender, RoutedEventArgs args) {
        if (DataContext is PageImportPreviewViewModel { CanApply: true }) Close(true);
    }
}
