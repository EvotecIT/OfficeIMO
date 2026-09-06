using Avalonia.Controls;
using Avalonia.Threading;
using System.ComponentModel;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Features.Editor;

public sealed partial class FormsInspectorView : UserControl {
    private MainWindowViewModel? _model;
    private bool _attached;

    public FormsInspectorView() {
        InitializeComponent();
        AttachedToVisualTree += (_, _) => { _attached = true; UpdateSubscription(); };
        DetachedFromVisualTree += (_, _) => { _attached = false; UpdateSubscription(); };
        DataContextChanged += (_, _) => UpdateSubscription();
    }

    private void UpdateSubscription() {
        if (_model is not null) _model.PropertyChanged -= OnModelChanged;
        _model = _attached ? DataContext as MainWindowViewModel : null;
        if (_model is not null) _model.PropertyChanged += OnModelChanged;
    }

    private void OnModelChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName != nameof(MainWindowViewModel.HasFormPreview) || _model?.HasFormPreview != true) return;
        Dispatcher.UIThread.Post(() => {
            if (_attached && _model?.HasFormPreview == true) FormAppearancePreview.BringIntoView();
        }, DispatcherPriority.Loaded);
    }
}
