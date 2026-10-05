using OfficeIMO.Studio.Infrastructure;
using Avalonia.Controls;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Editor;

public sealed partial class PdfSigningDialogContent : StudioDialogContent {
    public PdfSigningDialogContent() { InitializeComponent(); }
    internal PdfSigningDialogContent(PdfSigningPreviewViewModel model) : this() { DataContext = model; Opened += (_, _) => ReviewContent.Focus(); }
    private void OnCancel(object? sender, RoutedEventArgs args) => Close(false);
    private void OnApply(object? sender, RoutedEventArgs args) { if (DataContext is PdfSigningPreviewViewModel { IsPreview: true }) Close(true); }
}
