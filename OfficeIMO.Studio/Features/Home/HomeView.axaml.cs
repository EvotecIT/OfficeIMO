using Avalonia.Controls;
using Avalonia.Threading;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Features.Home;

public sealed partial class HomeView : UserControl {
    private static readonly TimeSpan PreviewDelay = TimeSpan.FromMilliseconds(900);
    private DispatcherTimer? _previewTimer;
    private CancellationTokenSource? _previewCancellation;
    private MainWindowViewModel? _previewContext;

    public HomeView() {
        InitializeComponent();
        OpenShortcut.Text = OperatingSystem.IsMacOS() ? "⌘O" : "Ctrl O";
        SizeChanged += (_, e) => ApplyResponsiveLayout(e.NewSize.Width);
        RecentList.ContainerPrepared += (_, _) => SchedulePreviews();
        DataContextChanged += (_, _) => AttachPreviewContext();
        AttachedToVisualTree += (_, _) => AttachPreviewContext();
        DetachedFromVisualTree += (_, _) => {
            CancelPreviews();
            if (_previewContext is not null) _previewContext.PropertyChanged -= OnPreviewContextPropertyChanged;
            _previewContext = null;
        };
        PropertyChanged += (_, change) => {
            if (change.Property != IsVisibleProperty) return;
            if (IsVisible) SchedulePreviews();
            else CancelPreviews();
        };
    }

    private void AttachPreviewContext() {
        if (_previewContext is not null) _previewContext.PropertyChanged -= OnPreviewContextPropertyChanged;
        _previewContext = DataContext as MainWindowViewModel;
        if (_previewContext is not null) _previewContext.PropertyChanged += OnPreviewContextPropertyChanged;
        CancelPreviews();
        if (IsEffectivelyVisible) SchedulePreviews();
    }

    private void OnPreviewContextPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs change) {
        if (change.PropertyName != nameof(MainWindowViewModel.CanStartDocumentTransition)) return;
        if (_previewContext?.CanStartDocumentTransition == false) CancelPreviews();
        else if (IsEffectivelyVisible) SchedulePreviews();
    }

    private void CancelPreviews() {
        _previewTimer?.Stop();
        _previewCancellation?.Cancel();
        _previewCancellation?.Dispose();
        _previewCancellation = null;
    }

    // Previews render only while Home is on screen and no document is opening, so they never
    // compete with the document the user is waiting for.
    private void SchedulePreviews() {
        if (!IsEffectivelyVisible || _previewContext?.CanStartDocumentTransition == false) return;
        _previewTimer ??= new DispatcherTimer(PreviewDelay, DispatcherPriority.Background, (_, _) => LoadPreviews());
        _previewTimer.Stop();
        _previewTimer.Start();
    }

    private void LoadPreviews() {
        _previewTimer?.Stop();
        if (!IsEffectivelyVisible) return;
        if (TopLevel.GetTopLevel(this) is MainWindow { IsStartingUp: true } ||
            DataContext is MainWindowViewModel { CanStartDocumentTransition: false }) {
            if (_previewContext?.CanStartDocumentTransition != false) SchedulePreviews();
            return;
        }
        _previewCancellation ??= new CancellationTokenSource();
        CancellationToken cancellationToken = _previewCancellation.Token;
        foreach (Control container in RecentList.GetRealizedContainers()) {
            if (container.DataContext is RecentDocumentViewModel recent) recent.EnsureThumbnail(cancellationToken);
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
