using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PdfOrganizerPageView : UserControl {
    // Load thumbnails slightly ahead of scrolling; the page grid is not virtualized, so
    // containers far outside the viewport keep their preview released.
    private const double PrefetchMargin = 480D;
    private PdfOrganizerPageViewModel? _viewModel;
    private bool _attached;
    private bool _nearViewport = true;

    public PdfOrganizerPageView() {
        InitializeComponent();
        DataContextChanged += (_, _) => UpdateViewModel();
        AttachedToVisualTree += (_, _) => {
            _attached = true;
            UpdateViewModel();
        };
        DetachedFromVisualTree += (_, _) => {
            _attached = false;
            _viewModel?.Detach();
        };
        EffectiveViewportChanged += OnEffectiveViewportChanged;
    }

    private void OnEffectiveViewportChanged(object? sender, EffectiveViewportChangedEventArgs e) {
        Rect viewport = e.EffectiveViewport;
        bool near = viewport.Width > 0 && viewport.Height > 0 &&
                    viewport.Inflate(PrefetchMargin).Intersects(new Rect(Bounds.Size));
        if (near == _nearViewport) return;
        _nearViewport = near;
        if (!_attached) return;
        if (near) _viewModel?.Attach();
        else _viewModel?.Detach();
    }

    private void UpdateViewModel() {
        if (ReferenceEquals(_viewModel, DataContext)) {
            if (_attached && _nearViewport) _viewModel?.Attach();
            return;
        }

        _viewModel?.Detach();
        _viewModel = DataContext as PdfOrganizerPageViewModel;
        if (_attached && _nearViewport) _viewModel?.Attach();
    }
}
