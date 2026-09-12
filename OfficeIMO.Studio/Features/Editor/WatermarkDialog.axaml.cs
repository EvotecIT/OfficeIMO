using Avalonia.Controls;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Editor;

public sealed partial class WatermarkDialog : Window {
    public WatermarkDialog() => InitializeComponent();

    public WatermarkDialog(WatermarkPreviewViewModel model) : this() {
        DataContext = model;
        Opened += async (_, _) => {
            if (Owner is { } owner) {
                Width = Math.Max(MinWidth, Math.Min(Width, owner.Bounds.Width - 40));
                Height = Math.Max(MinHeight, Math.Min(Height, owner.Bounds.Height - 40));
            }
            await model.PreviewCommand.ExecuteAsync(null);
        };
    }

    private void OnColor(object? sender, RoutedEventArgs e) {
        if (DataContext is WatermarkPreviewViewModel model && sender is Button { Tag: string color }) model.Color = color;
    }
    private void OnCancel(object? sender, RoutedEventArgs e) => Close(false);
    private void OnApply(object? sender, RoutedEventArgs e) {
        if (DataContext is WatermarkPreviewViewModel { CanApply: true }) Close(true);
    }
}
