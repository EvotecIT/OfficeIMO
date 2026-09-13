using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Assistant;

/// <summary>Owns setup presentation; the application owns the selected connection.</summary>
internal sealed partial class ConnectionsWindow : Window {
    public ConnectionsWindow() {
        InitializeComponent();
        Closed += (_, _) => { if (DataContext is StudioAiConnections connections) connections.CancelCommand.Execute(null); };
    }
    private void CloseClick(object? sender, RoutedEventArgs e) => Close(false);
    private void UseConnectionClick(object? sender, RoutedEventArgs e) {
        if (DataContext is StudioAiConnections connections && connections.RememberSelection()) Close(true);
    }
    protected override void OnKeyDown(KeyEventArgs e) {
        base.OnKeyDown(e);
        if (e.Key == Key.Escape) { e.Handled = true; Close(false); }
    }
}
