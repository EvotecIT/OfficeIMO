using Avalonia.Controls;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PageMoveDialog : Window {
    public PageMoveDialog() { InitializeComponent(); }
    internal PageMoveDialog(PageMovePreviewViewModel model) : this() {
        DataContext = model;
        Opened += (_, _) => DestinationInput.Focus();
    }
    private void OnCancel(object? sender, RoutedEventArgs args) => Close(false);
    private void OnApply(object? sender, RoutedEventArgs args) {
        if (DataContext is PageMovePreviewViewModel { CanApply: true }) Close(true);
    }
}
