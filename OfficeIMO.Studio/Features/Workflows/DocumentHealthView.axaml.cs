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
    }
}
