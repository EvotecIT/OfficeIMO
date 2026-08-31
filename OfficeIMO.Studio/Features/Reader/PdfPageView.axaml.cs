using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageView : UserControl {
    private PdfPageViewModel? _viewModel;
    private bool _attached;

    public PdfPageView() {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
        AttachedToVisualTree += (_, _) => {
            _attached = true;
            UpdateViewModel();
        };
        DetachedFromVisualTree += (_, _) => {
            _attached = false;
            _viewModel?.DetachFromViewport();
        };
    }

    private void OnDataContextChanged(object? sender, EventArgs e) {
        UpdateViewModel();
    }

    private void UpdateViewModel() {
        if (ReferenceEquals(_viewModel, DataContext)) {
            if (_attached) _viewModel?.AttachToViewport();
            return;
        }

        _viewModel?.DetachFromViewport();
        _viewModel = DataContext as PdfPageViewModel;
        if (_attached) _viewModel?.AttachToViewport();
    }
}
