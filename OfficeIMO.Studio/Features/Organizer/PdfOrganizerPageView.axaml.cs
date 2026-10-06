using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PdfOrganizerPageView : UserControl {
    // Load thumbnails slightly ahead of scrolling; the page grid is not virtualized, so
    // containers far outside the viewport keep their preview released.
    private const double PrefetchMargin = 480D;
    private PdfOrganizerPageViewModel? _viewModel;
    private TopLevel? _topLevel;
    private bool _attached;
    private bool _nearViewport = true;

    public PdfOrganizerPageView() {
        InitializeComponent();
        DataContextChanged += (_, _) => UpdateViewModel();
        AttachedToVisualTree += (_, _) => {
            _attached = true;
            TrackRenderScaling(TopLevel.GetTopLevel(this));
            UpdateViewModel();
        };
        DetachedFromVisualTree += (_, _) => {
            _attached = false;
            TrackRenderScaling(null);
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

    // Thumbnail bitmaps are requested in device pixels; moving between displays changes how many that is.
    private void TrackRenderScaling(TopLevel? topLevel) {
        if (ReferenceEquals(_topLevel, topLevel)) return;
        if (_topLevel is not null) _topLevel.ScalingChanged -= OnTopLevelScalingChanged;
        _topLevel = topLevel;
        if (_topLevel is not null) _topLevel.ScalingChanged += OnTopLevelScalingChanged;
    }

    private void OnTopLevelScalingChanged(object? sender, EventArgs e) => ApplyRenderScaling();

    private void ApplyRenderScaling() {
        if (_topLevel is not null) _viewModel?.SetRenderScaling(_topLevel.RenderScaling);
    }

    private void UpdateViewModel() {
        if (ReferenceEquals(_viewModel, DataContext)) {
            ApplyRenderScaling();
            if (_attached && _nearViewport) _viewModel?.Attach();
            return;
        }

        _viewModel?.Detach();
        _viewModel = DataContext as PdfOrganizerPageViewModel;
        ApplyRenderScaling();
        if (_attached && _nearViewport) _viewModel?.Attach();
    }
}
