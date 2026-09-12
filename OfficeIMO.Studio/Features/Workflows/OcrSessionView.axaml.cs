using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Workflows;

public partial class OcrSessionView : UserControl {
    public OcrSessionView() => InitializeComponent();

    private void OnEmptyStateSizeChanged(object? sender, SizeChangedEventArgs args) {
        if (sender is not Border panel) return;
        // The toolbar still offers Add files when expanded settings leave little vertical space.
        panel.Padding = new Avalonia.Thickness(args.NewSize.Height < 120 ? 4 : 24);
        EmptyDescription.IsVisible = args.NewSize.Height >= 200;
        EmptyAdd.IsVisible = args.NewSize.Height >= 120;
    }
}
