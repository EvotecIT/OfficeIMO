using OfficeIMO.Studio.Infrastructure;
using Avalonia.Controls;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Editor;

public sealed partial class WatermarkDialogContent : StudioDialogContent {
    private bool _compact;
    public WatermarkDialogContent() {
        InitializeComponent();
        SizeChanged += (_, e) => {
            bool compact = e.NewSize.Width < 720;
            if (_compact == compact) return;
            _compact = compact;
            if (compact) {
                WatermarkLayout.Children.Clear();
                OptionsTab.Content = OptionsPane;
                PreviewTab.Content = PreviewPane;
            } else {
                OptionsTab.Content = PreviewTab.Content = null;
                WatermarkLayout.Children.Add(OptionsPane);
                WatermarkLayout.Children.Add(PreviewPane);
            }
            CompactTabs.IsVisible = compact;
            WatermarkLayout.IsVisible = !compact;
        };
    }

    public WatermarkDialogContent(WatermarkPreviewViewModel model) : this() {
        DataContext = model;
        model.AutoPreview = true;
        Opened += async (_, _) => {
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
