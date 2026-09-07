using Avalonia.Controls;
using Avalonia.Threading;
using System.ComponentModel;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class OcrReviewView : UserControl {
    private OcrReviewViewModel? _observed;
    public OcrReviewView() {
        InitializeComponent();
        DataContextChanged += (_, _) => {
            if (_observed is not null) _observed.PropertyChanged -= ReviewChanged;
            _observed = DataContext as OcrReviewViewModel;
            if (_observed is not null) _observed.PropertyChanged += ReviewChanged;
        };
    }

    private void ReviewChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName is nameof(OcrReviewViewModel.SelectedWord) or nameof(OcrReviewViewModel.IsZoomed) or nameof(OcrReviewViewModel.Preview))
            Dispatcher.UIThread.Post(() => {
                if (_observed?.IsZoomed == true && _observed.HasSelectedWord) ZoomSelection.BringIntoView();
            }, DispatcherPriority.Loaded);
    }
}
