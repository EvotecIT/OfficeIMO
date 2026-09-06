using Avalonia.Controls;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PageSplitDialog : Window {
    public PageSplitDialog() { InitializeComponent(); }
    internal PageSplitDialog(PageSplitPreviewViewModel model) : this() {
        DataContext = model;
        Opened += (_, _) => { if (model.IsPreview) PartSizeInput.Focus(); else if (model.Files.Count > 0) ResultFiles.SelectedIndex = 0; };
    }
    private void OnCancel(object? sender, RoutedEventArgs args) => Close(false);
    private void OnApply(object? sender, RoutedEventArgs args) {
        if (DataContext is PageSplitPreviewViewModel { CanApply: true }) Close(true);
    }
}
