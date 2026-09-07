using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class SearchablePdfOcrView : UserControl {
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
