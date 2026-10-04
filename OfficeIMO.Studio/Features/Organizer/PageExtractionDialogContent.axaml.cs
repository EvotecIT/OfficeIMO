using OfficeIMO.Studio.Infrastructure;
using Avalonia.Controls;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PageExtractionDialogContent : StudioDialogContent {
    public PageExtractionDialogContent() { InitializeComponent(); }
    internal PageExtractionDialogContent(PageExtractionPreviewViewModel model) : this() {
        DataContext = model;
        Opened += (_, _) => { if (model.IsPreview) PageRangeInput.Focus(); else PageOrder.Focus(); };
    }
    private void OnCancel(object? sender, RoutedEventArgs args) => Close(false);
    private void OnApply(object? sender, RoutedEventArgs args) {
        if (DataContext is PageExtractionPreviewViewModel { CanApply: true }) Close(true);
    }
}
