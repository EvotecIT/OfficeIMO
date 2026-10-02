using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Home;

public sealed partial class ToolsView : UserControl {
    public ToolsView() {
        InitializeComponent();
        SizeChanged += (_, e) => ToolsContent.Margin = e.NewSize.Width < 600
            ? new Avalonia.Thickness(16, 20, 16, 24) : new Avalonia.Thickness(36, 28, 36, 36);
    }
}
