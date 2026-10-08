using System.ComponentModel;
using Avalonia.Automation;
using Avalonia.Interactivity;
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
    internal void FocusActiveIndependentPane() => IndependentPanes.FocusActivePane();
    private void OnIndependentPanesChanged(object? sender, PropertyChangedEventArgs args) => ShowContextPanes();
    private void OnIndependentPanesClick(object? sender, RoutedEventArgs args) {
        Panes?.SplitCommand.Execute(null);
        FocusActiveIndependentPane();
    }
}
