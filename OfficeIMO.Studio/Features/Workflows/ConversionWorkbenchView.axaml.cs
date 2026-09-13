using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionWorkbenchView : UserControl {
    public ConversionWorkbenchView() {
        InitializeComponent();
        SizeChanged += OnSizeChanged;
    }

    internal bool IsCompactLayout { get; private set; }

    private void OnSizeChanged(object? sender, SizeChangedEventArgs e) => ApplyResponsiveLayout(e.NewSize.Width);

    internal void ApplyResponsiveLayout(double width) {
        IsCompactLayout = width < 1000D;
        WorkspaceGrid.ColumnDefinitions[0].Width = new GridLength(IsCompactLayout ? 238D : Math.Clamp(width * .2D, 270D, 360D));
        WorkspaceGrid.ColumnDefinitions[2].Width = new GridLength(IsCompactLayout ? 0D : Math.Clamp(width * .25D, 320D, 480D));
        Grid.SetRow(DetailsPanel, IsCompactLayout ? 1 : 0);
        Grid.SetColumn(DetailsPanel, IsCompactLayout ? 0 : 2);
        Grid.SetColumnSpan(DetailsPanel, IsCompactLayout ? 3 : 1);
        DetailsPanel.MaxHeight = IsCompactLayout ? 230D : double.PositiveInfinity;
    }
}
