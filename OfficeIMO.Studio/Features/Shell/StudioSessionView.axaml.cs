using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class StudioSessionView : UserControl {
    public StudioSessionView() => InitializeComponent();

    // Hides the restore banner for this run only; the session stays available until the user forgets it.
    private void OnDismissClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => IsVisible = false;
}
