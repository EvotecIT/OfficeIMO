using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>In-window command search over the shared <see cref="StudioCommandCatalog"/>.</summary>
internal sealed partial class StudioCommandPalette : UserControl {
    private TaskCompletionSource<StudioCommandItem?>? _completion;
    private StudioCommandCatalog? _catalog;

    public StudioCommandPalette() {
        InitializeComponent();
        AddHandler(KeyDownEvent, OnPaletteKeyDown, RoutingStrategies.Tunnel);
        SizeChanged += (_, _) => UpdatePresentation();
        DetachedFromVisualTree += (_, _) => Dismiss();
    }

    internal StudioCommandPaletteModel? Model { get; private set; }

    internal bool IsOpen => _completion is not null;

    /// <summary>Adapts the shared command search for touch; the host supplies the keyboard-safe bounds.</summary>
    internal void UseTouchPresentation() {
        Classes.Add("touchPalette");
        UpdatePresentation();
    }

    private void UpdatePresentation() {
        bool touch = Classes.Contains("touchPalette");
        bool compact = Bounds.Width < 600;
        Classes.Set("compactPalette", compact);
        QueryBox.MinHeight = touch ? 44 : 38;
        RunButton.MinHeight = touch ? 44 : 30;
        PalettePanel.Margin = touch ? new Thickness(12) : new Thickness(16, 36, 16, 16);
        PalettePanel.MaxHeight = Math.Max(1, Math.Min(520, Bounds.Height - PalettePanel.Margin.Top - PalettePanel.Margin.Bottom));
        KeyboardHelp.IsVisible = !touch;
    }

    /// <summary>Shows the palette and completes with the chosen available command, or null when dismissed.</summary>
    internal Task<StudioCommandItem?> ShowAsync(StudioCommandCatalog catalog) {
        if (_completion is not null) return _completion.Task;
        _catalog = catalog;
        Model = new StudioCommandPaletteModel(catalog);
        DataContext = Model;
        _completion = new TaskCompletionSource<StudioCommandItem?>();
        IsVisible = true;
        this.Dispatcher.Post(() => {
            if (!IsOpen) return;
            QueryBox.Focus();
            QueryBox.SelectAll();
        }, DispatcherPriority.Loaded);
        return _completion.Task;
    }

    internal void Dismiss() => Complete(null);

    private void Complete(StudioCommandItem? command) {
        if (_completion is not { } completion) return;
        _completion = null;
        IsVisible = false;
        if (command is not null) _catalog?.MarkUsed(command.Id);
        DataContext = null;
        Model = null;
        completion.TrySetResult(command);
    }

    private void OnRunClick(object? sender, RoutedEventArgs e) => SelectCommand();
    private void OnCloseClick(object? sender, RoutedEventArgs e) => Dismiss();

    private void OnResultTapped(object? sender, TappedEventArgs e) {
        if (e.Source is Control { DataContext: StudioCommandItem }) SelectCommand();
    }

    private void OnScrimPressed(object? sender, PointerPressedEventArgs e) {
        Dismiss();
        e.Handled = true;
    }

    private void SelectCommand() {
        if (Model?.SelectedCommand is { IsAvailable: true } command) Complete(command);
    }

    private void OnPaletteKeyDown(object? sender, KeyEventArgs e) {
        if (Model is null) return;
        if (e.Key == Key.Escape) { Dismiss(); e.Handled = true; }
        else if (e.Key == Key.Enter) {
            if (CloseButton.IsKeyboardFocusWithin) Dismiss(); else SelectCommand();
            e.Handled = true;
        }
        else if (e.Key is Key.Down or Key.Up) {
            int next = Math.Clamp(ResultsList.SelectedIndex + (e.Key == Key.Down ? 1 : -1), 0, Math.Max(0, Model.Results.Count - 1));
            ResultsList.SelectedIndex = next;
            if (ResultsList.SelectedItem is { } item) ResultsList.ScrollIntoView(item);
            e.Handled = true;
        }
    }
}
