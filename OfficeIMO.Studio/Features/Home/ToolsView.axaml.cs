using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Home;

public sealed partial class ToolsView : UserControl {
    public ToolsView() {
        InitializeComponent();
        SizeChanged += (_, e) => ToolsContent.Classes.Set("compact", e.NewSize.Width < 600);
    }
}
