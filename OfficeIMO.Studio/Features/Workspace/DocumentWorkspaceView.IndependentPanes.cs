using System.ComponentModel;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Features.Workspace;

public sealed partial class DocumentWorkspaceView {
    private StudioDocumentPanes? _panes;
    internal StudioDocumentPanes? Panes {
        get => _panes;
        set {
            if (_panes is not null) _panes.PropertyChanged -= OnIndependentPanesChanged;
            _panes = value;
            IndependentPanes.DataContext = value;
            if (value is not null) {
                value.PropertyChanged += OnIndependentPanesChanged;
                IndependentPanesLabel.Text = value.SplitLabel;
                AutomationProperties.SetName(IndependentPanesButton, value.SplitLabel);
                Avalonia.Controls.ToolTip.SetTip(IndependentPanesButton, value.HelpLabel);
            }
            ShowContextPanes();
        }
    }
    internal void FocusActiveIndependentPane() {
        var requestedPane = Panes?.ActivePane;
        if (requestedPane is null) return;
        if (_phoneLayout) SetPanes(false, false);
        IndependentPanes.FocusActivePane();
        var focusManager = TopLevel.GetTopLevel(this)?.FocusManager;
        var initialFocus = focusManager?.GetFocusedElement();
        // Changing document tabs can replace compact context controls after the first focus attempt.
        Dispatcher.Post(() => {
            if (Panes?.ActivePane != requestedPane || _document?.IsPdfWorkspaceMode != true ||
                !IndependentPanes.IsEffectivelyVisible) return;
            var currentFocus = focusManager?.GetFocusedElement();
            if (currentFocus is not null && !ReferenceEquals(currentFocus, initialFocus)) return;
            if (_phoneLayout) SetPanes(false, false);
            IndependentPanes.FocusActivePane();
        }, DispatcherPriority.Loaded);
    }
    private void OnIndependentPanesChanged(object? sender, PropertyChangedEventArgs args) => ShowContextPanes();
    private void OnIndependentPanesClick(object? sender, RoutedEventArgs args) {
        Panes?.SplitCommand.Execute(null);
        FocusActiveIndependentPane();
    }
}
