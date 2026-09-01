using Avalonia.Controls;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageView : UserControl {
    private PdfPageViewModel? _viewModel;
    private bool _attached;

    public PdfPageView() {
        InitializeComponent();
        PageCanvas.LinkActivated += OnLinkActivated;
        PageCanvas.EditorGestureCompleted += OnEditorGestureCompleted;
        PageCanvas.ObjectSelected += OnObjectSelected;
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

    private void OnLinkActivated(string target) => _viewModel?.ActivateLink(target);

    private void OnEditorGestureCompleted(PdfEditorGesture gesture) => _viewModel?.CompleteEditorGesture(gesture);

    private void OnObjectSelected(PdfEditorSelection? selection) => _viewModel?.SelectObject(selection);

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
