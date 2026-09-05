using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Organizer;

public sealed partial class PdfOrganizerPageView : UserControl {
    private PdfOrganizerPageViewModel? _viewModel;
    private bool _attached;

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
    }

    private void UpdateViewModel() {
        if (ReferenceEquals(_viewModel, DataContext)) {
            if (_attached) _viewModel?.Attach();
            return;
        }

        _viewModel?.Detach();
        _viewModel = DataContext as PdfOrganizerPageViewModel;
        if (_attached) _viewModel?.Attach();
    }
}
