using OfficeIMO.Studio.Infrastructure;
using Avalonia.Controls;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PageMoveDialogContent : StudioDialogContent {
    public PageMoveDialogContent() { InitializeComponent(); }
    internal PageMoveDialogContent(PageMovePreviewViewModel model) : this() {
        DataContext = model;
        Opened += (_, _) => DestinationInput.Focus();
    }
    private void OnCancel(object? sender, RoutedEventArgs args) => Close(false);
    private void OnApply(object? sender, RoutedEventArgs args) {
        if (DataContext is PageMovePreviewViewModel { CanApply: true }) Close(true);
    }
}
