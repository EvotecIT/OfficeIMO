using Avalonia.Controls;
using Avalonia.Threading;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Features.Home;

public sealed partial class HomeView : UserControl {
    private static readonly TimeSpan PreviewDelay = TimeSpan.FromMilliseconds(900);
    private DispatcherTimer? _previewTimer;

    public HomeView() {
        InitializeComponent();
        OpenShortcut.Text = OperatingSystem.IsMacOS() ? "⌘O" : "Ctrl O";
        SizeChanged += (_, e) => ApplyResponsiveLayout(e.NewSize.Width);
        RecentList.ContainerPrepared += (_, _) => SchedulePreviews();
        PropertyChanged += (_, change) => {
            if (change.Property == IsVisibleProperty && IsVisible) SchedulePreviews();
        };
    }

    // Previews render only while Home is on screen and no document is opening, so they never
    // compete with the document the user is waiting for.
    private void SchedulePreviews() {
        _previewTimer ??= new DispatcherTimer(PreviewDelay, DispatcherPriority.Background, (_, _) => LoadPreviews());
        _previewTimer.Stop();
        _previewTimer.Start();
    }

    private void LoadPreviews() {
        _previewTimer?.Stop();
        if (!IsEffectivelyVisible) return;
        if (TopLevel.GetTopLevel(this) is MainWindow { IsStartingUp: true } ||
            DataContext is MainWindowViewModel { CanStartDocumentTransition: false }) {
            SchedulePreviews();
            return;
        }
        foreach (Control container in RecentList.GetRealizedContainers()) {
            if (container.DataContext is RecentDocumentViewModel recent) recent.EnsureThumbnail();
        }
    }

    // The drop target and quick tasks sit side by side on wide windows and stack on narrow ones.
    private void ApplyResponsiveLayout(double width) {
        bool stacked = width < 1000D;
        Grid.SetColumn(QuickTasks, stacked ? 0 : 1);
        Grid.SetRow(QuickTasks, stacked ? 1 : 0);
        Grid.SetColumnSpan(DropTarget, stacked ? 2 : 1);
        Grid.SetColumnSpan(QuickTasks, stacked ? 2 : 1);
    }
}
