using Avalonia;
using Avalonia.Controls;
using Avalonia.Automation;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Input.Raw;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed partial class StudioMobileApplicationTests {
    [Theory]
    [InlineData(1366, 1024)]
    [InlineData(390, 844)]
    [InlineData(844, 390)]
    public async Task CommandSearchRoutesSharedActionsAndProtectsModalFocus(int width, int height) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-commands-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                var view = new MobileWorkspaceView();
                using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(null), _ => Task.CompletedTask);
                view.Connect(controller);
                var window = new Window { Content = view, Width = width, Height = height };
                try {
                    window.Show();
                    await controller.OpenSampleAsync();
                    var document = controller.Document;
                    Layout(window);
                    var navigation = view.FindControl<SplitView>("ApplicationNavigation")!;
                    var palette = view.FindControl<StudioCommandPalette>("CommandPalette")!;
                    var query = palette.FindControl<TextBox>("QueryBox")!;
                    var primary = OperatingSystem.IsMacOS() ? RawInputModifiers.Meta : RawInputModifiers.Control;
                    navigation.IsPaneOpen = true;
                    Layout(window);
                    var jobs = view.GetVisualDescendants().OfType<Button>()
                        .Single(button => AutomationProperties.GetName(button) == "Jobs");
                    jobs.BringIntoView();
                    Layout(window);
                    await WaitForCommandButtonAsync(window, jobs);
                    Click(window, jobs);
                    Layout(window);
                    Assert.True(document.IsJobsMode);
                    view.FindControl<Button>("ApplicationMenuButton")!.Focus();
                    window.KeyPress(Key.K, primary, PhysicalKey.None, null);
                    Layout(window);
                    Assert.True(palette.IsOpen);
                    Assert.False(navigation.IsEnabled);
                    Assert.Same(query, window.FocusManager!.GetFocusedElement());
                    Assert.True(query.Bounds.Height >= 44);
                    query.Text = "no-such-command";
                    Layout(window);
                    Assert.False(palette.FindControl<Button>("RunButton")!.IsEffectivelyEnabled);
                    query.Text = "Organize pages";
                    Layout(window);
                    Assert.Equal("Pages", palette.Model!.SelectedCommand!.Id);
                    Assert.True(palette.FindControl<ListBox>("ResultsList")!.ContainerFromIndex(0)!.Bounds.Height >= 56);
                    Assert.All(palette.GetVisualDescendants().OfType<StackPanel>().Where(panel => panel.Classes.Contains("commandMetadata")),
                        metadata => Assert.Equal(width >= 600, metadata.IsEffectivelyVisible));
                    Capture(window, $"command-search-{width}");
                    window.KeyPress(Key.W, primary, PhysicalKey.None, null);
                    Assert.Single(controller.Tabs.Tabs);
                    window.KeyPress(Key.Tab, RawInputModifiers.None, PhysicalKey.None, null);
                    Assert.Same(palette.FindControl<Button>("CloseButton"), window.FocusManager.GetFocusedElement());
                    window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                    Layout(window);
                    Assert.False(palette.IsOpen);
                    Assert.True(document.IsJobsMode);
                    window.KeyPress(Key.K, primary, PhysicalKey.None, null);
                    Layout(window);
                    query.Text = "Organize pages";
                    Layout(window);
                    window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                    Layout(window);
                    Assert.False(palette.IsOpen);
                    Assert.True(navigation.IsEnabled);
                    Assert.True(document.IsPagesDocumentMode);

                    // A hardware-keyboard search from a specialist workspace returns to this document.
                    document.Commands["Settings"].Execute(null);
                    Layout(window);
                    view.FindControl<Button>("ApplicationMenuButton")!.Focus();
                    window.KeyPress(Key.F, primary, PhysicalKey.None, null);
                    Layout(window);
                    Assert.True(document.IsPdfWorkspaceMode);

                    document.ShowViewModeCommand.Execute(null);
                    document.Commands["FocusReading"].Execute(null);
                    Layout(window);
                    view.FindControl<Button>("ApplicationMenuButton")!.Focus();
                    window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                    Layout(window);
                    Assert.False(document.IsFocusReading);
                    window.KeyPress(Key.D1, primary, PhysicalKey.None, null);
                    Assert.Equal(1, document.Zoom);

                    // Touch entry and explicit dismissal remain reachable without a keyboard.
                    navigation.IsPaneOpen = true;
                    Layout(window);
                    view.FindControl<Button>("CommandSearchButton")!.BringIntoView();
                    Layout(window);
                    await WaitForCommandButtonAsync(window, view.FindControl<Button>("CommandSearchButton")!);
                    Click(window, view.FindControl<Button>("CommandSearchButton")!);
                    Layout(window);
                    Assert.True(palette.IsOpen);
                    Click(window, palette.FindControl<Button>("CloseButton")!);
                    Layout(window);
                    Assert.False(palette.IsOpen);
                    Assert.True(navigation.IsEnabled);
                    Assert.True(((Control)window.FocusManager.GetFocusedElement()!).IsEffectivelyVisible);

                    navigation.IsPaneOpen = true;
                    Layout(window);
                    Click(window, view.FindControl<Button>("CommandSearchButton")!);
                    Layout(window);
                    window.Close();
                    Assert.False(palette.IsOpen);
                } finally { window.Close(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    private static async Task WaitForCommandButtonAsync(Window window, Button button) {
        await StudioHeadlessInput.WaitForTargetAsync(window, button, () => Layout(window));
    }
}
