using Avalonia.Controls;
using Avalonia.Interactivity;
namespace OfficeIMO.Studio.Features.Assistant;
public sealed partial class AssistantView : UserControl {
    internal Func<StudioAiConnections, Task<bool>>? ManageConnectionsAsync { get; set; }

    public AssistantView() => InitializeComponent();

    private async void ManageConnectionsClick(object? sender, RoutedEventArgs e) {
        if (DataContext is not DocumentAssistantViewModel model) return;
        bool selected = ManageConnectionsAsync is { } manage
            ? await manage(model.Connections)
            : TopLevel.GetTopLevel(this) is Window owner && await new ConnectionsWindow { DataContext = model.Connections }.ShowDialog<bool>(owner);
        if (selected && ReferenceEquals(DataContext, model)) model.ShowConnections = false;
    }
}
