using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed partial class StudioCommandPalette : Window {
    internal StudioCommandPalette(StudioCommandCatalog catalog) {
        InitializeComponent();
        Model = new StudioCommandPaletteModel(catalog);
        DataContext = Model;
        Opened += (_, _) => QueryBox.Focus();
        KeyDown += OnPaletteKeyDown;
    }

    internal StudioCommandPaletteModel Model { get; }

    private void OnRunClick(object? sender, RoutedEventArgs e) => SelectCommand();
    private void OnResultDoubleTapped(object? sender, TappedEventArgs e) => SelectCommand();

    private void SelectCommand() {
        if (Model.SelectedCommand is { IsAvailable: true } command) Close(command);
    }

    private void OnPaletteKeyDown(object? sender, KeyEventArgs e) {
        if (e.Key == Key.Escape) { Close(); e.Handled = true; }
        else if (e.Key == Key.Enter) { SelectCommand(); e.Handled = true; }
        else if (e.Key is Key.Down or Key.Up && ReferenceEquals(TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement(), QueryBox)) {
            int next = Math.Clamp(ResultsList.SelectedIndex + (e.Key == Key.Down ? 1 : -1), 0, Math.Max(0, Model.Results.Count - 1));
            ResultsList.SelectedIndex = next;
            if (ResultsList.SelectedItem is { } item) ResultsList.ScrollIntoView(item);
            e.Handled = true;
        }
    }
}
