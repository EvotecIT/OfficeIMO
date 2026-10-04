using OfficeIMO.Studio.Infrastructure;
using Avalonia.Controls;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Editor;

public sealed partial class PdfProtectionDialogContent : StudioDialogContent {
    public PdfProtectionDialogContent() { InitializeComponent(); }
    internal PdfProtectionDialogContent(PdfProtectionPreviewViewModel model) : this() {
        DataContext = model; Opened += (_, _) => ReviewContent.Focus();
    }
    private void OnCancel(object? sender, RoutedEventArgs args) => Close(false);
    private void OnApply(object? sender, RoutedEventArgs args) {
        if (DataContext is PdfProtectionPreviewViewModel { IsPreview: true }) Close(true);
    }
}
