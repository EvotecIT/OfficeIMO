using System.ComponentModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Features.Workspace;

/// <summary>Shows two independent viewports, or the active viewport with a keyboard-accessible switch on compact windows.</summary>
public sealed partial class IndependentDocumentPanesView : UserControl {
    private StudioDocumentPanes? _panes;
    private bool _compact;
    private bool _attached;
    private bool _observing;
    public IndependentDocumentPanesView() {
        InitializeComponent();
        SizeChanged += (_, args) => ApplyLayout(args.NewSize.Width);
        DataContextChanged += (_, _) => {
            StopObserving();
            _panes = DataContext as StudioDocumentPanes;
            Observe();
            ApplyLayout(Bounds.Width);
        };
        AttachedToVisualTree += (_, _) => {
            _attached = true;
            Observe();
            ApplyLayout(Bounds.Width);
        };
        DetachedFromVisualTree += (_, _) => {
            _attached = false;
            StopObserving();
        };
    }
    internal bool IsCompact => _compact;
    internal void FocusActivePane() => (_panes?.ActivePane == _panes?.Right ? RightPane : LeftPane).FocusReader();
    private void OnPanesChanged(object? sender, PropertyChangedEventArgs args) => ApplyLayout(Bounds.Width);
    private void Observe() {
        if (!_attached || _observing || _panes is null) return;
        _panes.PropertyChanged += OnPanesChanged;
        _observing = true;
    }
    private void StopObserving() {
        if (!_observing) return;
        if (_panes is not null) _panes.PropertyChanged -= OnPanesChanged;
        _observing = false;
    }
    private void OnSwitchClick(object? sender, RoutedEventArgs args) { _panes?.SwitchPane(); FocusActivePane(); }
    private void ApplyLayout(double width) {
        _compact = width < 780;
        PaneSplitter.IsVisible = !_compact;
        PaneGrid.ColumnDefinitions[0].Width = new GridLength(1, GridUnitType.Star);
        PaneGrid.ColumnDefinitions[1].Width = new GridLength(_compact ? 0 : 6);
        PaneGrid.ColumnDefinitions[2].Width = _compact ? new GridLength(0) : new GridLength(1, GridUnitType.Star);
        Grid.SetColumn(RightPane, _compact ? 0 : 2);
        LeftPane.IsVisible = !_compact || _panes?.ActivePane == _panes?.Left;
        RightPane.IsVisible = !_compact || _panes?.ActivePane == _panes?.Right;
    }
}
