using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class DocumentHealthView : UserControl {
    public DocumentHealthView() {
        InitializeComponent();
        SizeChanged += OnSizeChanged;
    }

    internal bool IsCompactLayout { get; private set; }

    private void OnSizeChanged(object? sender, SizeChangedEventArgs e) => ApplyResponsiveLayout(e.NewSize.Width);

    internal void ApplyResponsiveLayout(double width) {
        IsCompactLayout = width < 1000D;
        WorkspaceGrid.ColumnDefinitions[0].Width = new GridLength(IsCompactLayout ? 250D : 286D);
        WorkspaceGrid.ColumnDefinitions[2].Width = new GridLength(IsCompactLayout ? 0D : 320D);
        DetailsPanel.IsVisible = !IsCompactLayout;
    }
}
