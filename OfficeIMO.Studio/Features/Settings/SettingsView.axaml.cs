using Avalonia.Controls;

namespace OfficeIMO.Studio.Features.Settings;

public sealed partial class SettingsView : UserControl {
    public SettingsView() {
        InitializeComponent();
        FocusReadingHelp.IsVisible = !OperatingSystem.IsIOS();
        SizeChanged += (_, e) => {
            bool compact = e.NewSize.Width < 640;
            foreach (var grid in new[] { AppearanceGrid, PrivacyGrid }) {
                grid.ColumnDefinitions = new ColumnDefinitions(compact ? "*" : "*,*");
                grid.RowDefinitions = new RowDefinitions(compact ? "Auto,Auto" : "Auto");
                Grid.SetColumn(grid.Children[1], compact ? 0 : 1);
                Grid.SetRow(grid.Children[1], compact ? 1 : 0);
            }
        };
    }

    internal void UseTouchPresentation() => FocusReadingHelp.IsVisible = false;
}
