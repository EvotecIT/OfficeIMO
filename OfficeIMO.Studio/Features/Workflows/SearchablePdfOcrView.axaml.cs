using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class SearchablePdfOcrView : UserControl {
    private async void CopyExtractedText(object? sender, RoutedEventArgs args) {
        if (DataContext is not MainWindowViewModel shell || !shell.OcrWorkbench.HasExtractedText) return;
        var model = shell.OcrWorkbench;
        try {
            var clipboard = TopLevel.GetTopLevel(this)?.Clipboard
                ?? throw new InvalidOperationException(StudioLocalization.Current.GetOrDefault("Ocr.Text.NoClipboard", "Clipboard is unavailable. Select the text to copy it manually."));
            await clipboard.SetTextAsync(model.ExtractedText);
            model.Status = StudioLocalization.Current.GetOrDefault("Ocr.Text.Copied", "Reviewed text copied.");
        } catch (Exception error) { model.ErrorMessage = error.Message; }
    }

    public SearchablePdfOcrView() {
        InitializeComponent();
        SizeChanged += (_, e) => {
            bool compact = e.NewSize.Width < 1000D;
            OcrColumns.ColumnDefinitions[0].Width = compact ? GridLength.Star : new GridLength(400D);
            OcrColumns.ColumnDefinitions[1].Width = compact ? new GridLength(0D) : GridLength.Star;
            Grid.SetColumn(InformationPanel, compact ? 0 : 1);
            Grid.SetRow(InformationPanel, compact ? 1 : 0);
        };
    }
}
