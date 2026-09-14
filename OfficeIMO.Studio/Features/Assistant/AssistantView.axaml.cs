using Avalonia.Controls;
using Avalonia.Interactivity;
namespace OfficeIMO.Studio.Features.Assistant;
public sealed partial class AssistantView : UserControl {
    public AssistantView() => InitializeComponent();

    private async void ManageConnectionsClick(object? sender, RoutedEventArgs e) {
        if (DataContext is not DocumentAssistantViewModel model || TopLevel.GetTopLevel(this) is not Window owner) return;
        var window = new ConnectionsWindow { DataContext = model.Connections };
        if (await window.ShowDialog<bool>(owner) && ReferenceEquals(DataContext, model)) model.ShowConnections = false;
    }
}
