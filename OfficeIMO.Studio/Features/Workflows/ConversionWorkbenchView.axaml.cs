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
        WorkspaceGrid.ColumnDefinitions[0].Width = new GridLength(IsCompactLayout ? 238D : 270D);
        WorkspaceGrid.ColumnDefinitions[2].Width = new GridLength(IsCompactLayout ? 0D : 300D);
        DetailsPanel.IsVisible = !IsCompactLayout;
    }
}
