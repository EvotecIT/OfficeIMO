using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class StudioJobsView : UserControl {
    public StudioJobsView() {
        InitializeComponent();
        // The starting actions navigate the desktop window; touch hosts offer their own navigation.
        AttachedToVisualTree += (_, _) => EmptyActions.IsVisible = TopLevel.GetTopLevel(this) is Window;
    }
}