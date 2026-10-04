using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless;
using Avalonia.Interactivity;
using Avalonia.Styling;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioAdaptiveShellTests {
    [Theory]
    [InlineData(390, 844, false)]
    [InlineData(700, 500, true)]
    [InlineData(834, 1112, false)]
    [InlineData(1280, 820, true)]
    public async Task NavigationAndDocumentActionsRemainReachableWhenResizing(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var window = new MainWindow { RequestedThemeVariant = dark ? ThemeVariant.Dark : ThemeVariant.Light };
            try {
                window.Show();
                Layout(window, width, height);
                Capture(window, $"shell-{width}-home");
                var navigation = window.FindControl<Control>(width < 700 ? "CompactNavigation" : "NavigationRail")!;
                Assert.True(navigation.IsEffectivelyVisible);
                foreach (var button in navigation.GetVisualDescendants().OfType<Button>().Where(button => button.IsEffectivelyVisible))
                    AssertInside(window, button);
                AssertInside(window, window.FindControl<Button>("OpenDocumentTabButton")!);
                AssertInside(window, window.FindControl<Button>("CommandSearchButton")!);

                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                Layout(window, width, height);
                var workspace = window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!;
                foreach (string name in new[] { "DocumentModePicker", "DocumentMenuButton", "SaveButton", "InspectorToggle", "NavigationToggle" }) {
                    var control = workspace.FindControl<Control>(name)!;
                    if (control.IsEffectivelyVisible) AssertInside(window, control);
                }
                Capture(window, $"shell-{width}-document");
                if (width < 700) {
                    AssertInside(window, window.FindControl<ComboBox>("CompactDocumentPicker")!);
                    double canvasWidth = window.ReaderPagesListControl.Bounds.Width;
                    var toggle = workspace.FindControl<ToggleButton>("InspectorToggle")!;
                    toggle.IsChecked = true;
                    toggle.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                    Layout(window, width, height);
                    Assert.True(workspace.FindControl<Grid>("InspectorPane")!.IsEffectivelyVisible);
                    Assert.Equal(canvasWidth, window.ReaderPagesListControl.Bounds.Width);
                    Capture(window, $"shell-{width}-inspector");
                    // Opening search replaces the inspector overlay without squeezing the document.
                    workspace.FocusSearch();
                    Layout(window, width, height);
                    Assert.False(workspace.FindControl<Grid>("InspectorPane")!.IsVisible);
                    AssertInside(window, workspace.FindControl<TextBox>("SearchBox")!);
                    Assert.Equal(canvasWidth, window.ReaderPagesListControl.Bounds.Width);
                }
                Layout(window, 1280, 820);
                Assert.False(window.FindControl<Control>("CompactNavigation")!.IsVisible);
                Assert.True(window.FindControl<TabControl>("DocumentTabs")!.IsEffectivelyVisible);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task RestoredCompactPaneHasUsableWidthOnFirstLayout(bool navigation) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var window = new MainWindow { Width = 390, Height = 844 };
            try {
                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                window.ViewModel.UpdatePanePreferences(238, 300, new(navigation, !navigation));
                window.Show();
                Layout(window, 390, 844);
                var workspace = window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!;
                var pane = workspace.FindControl<Grid>(navigation ? "NavigationPane" : "InspectorPane")!;
                AssertInside(window, pane);
                Assert.InRange(pane.Bounds.Width, 300, 340);
                Assert.True(window.ReaderPagesListControl.Bounds.Width >= workspace.Bounds.Width - 30);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task CompactHeaderReservesCaptionSpaceWithoutCoveringActions() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var window = new MainWindow();
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                // Exercise the shared header with the caption reservation used by Windows chrome.
                window.FindControl<Border>("CaptionButtonsSpacer")!.Width = 138;
                Layout(window, 390, 844);
                var picker = window.FindControl<ComboBox>("CompactDocumentPicker")!;
                AssertInside(window, picker);
                var pickerBounds = new Rect(picker.TranslatePoint(default, window)!.Value, picker.Bounds.Size);
                foreach (string name in new[] { "CommandSearchButton", "OpenDocumentTabButton", "CaptionButtonsSpacer" }) {
                    var control = window.FindControl<Control>(name)!;
                    AssertInside(window, control);
                    var bounds = new Rect(control.TranslatePoint(default, window)!.Value, control.Bounds.Size);
                    Assert.False(pickerBounds.Intersects(bounds), $"Picker {pickerBounds} overlaps {name}: {bounds}");
                }
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task NativeFileMenuTracksTheSelectedDocumentAndItsGuards() {
        if (!OperatingSystem.IsMacOS()) return;
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var window = new MainWindow();
            try {
                window.Show();
                var rootMenu = NativeMenu.GetMenu(window)!;
                var file = rootMenu.Items.OfType<NativeMenuItem>().First().Menu!;
                var save = file.Items.OfType<NativeMenuItem>().Single(item => ReferenceEquals(item.Command, window.ViewModel.Commands["Save"]));
                Assert.False(save.Command!.CanExecute(null));
                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                var document = window.ViewModel;
                Assert.Equal("openpreserve-pdfa1b-text.pdf", window.Title);
                file = NativeMenu.GetMenu(window)!.Items.OfType<NativeMenuItem>().First().Menu!;
                Assert.Contains(file.Items.OfType<NativeMenuItem>(), item => ReferenceEquals(item.Command, document.Commands["SaveAs"]) && item.Command.CanExecute(null));
                await window.TabHost.CloseSelectedTabAsync();
                Assert.Same(rootMenu, NativeMenu.GetMenu(window));
                file = NativeMenu.GetMenu(window)!.Items.OfType<NativeMenuItem>().First().Menu!;
                Assert.DoesNotContain(file.Items.OfType<NativeMenuItem>(), item => ReferenceEquals(item.Command, document.Commands["SaveAs"]));
                Assert.Equal("OfficeIMO Studio", window.Title);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void AssertInside(Window window, Control control) {
        Assert.True(control.IsEffectivelyVisible);
        Point origin = control.TranslatePoint(default, window)!.Value;
        Assert.True(origin.X >= -1 && origin.Y >= -1 && origin.X + control.Bounds.Width <= window.Bounds.Width + 1
            && origin.Y + control.Bounds.Height <= window.Bounds.Height + 1, $"{control.Name}: {origin} {control.Bounds} in {window.Bounds}");
    }

    private static void Layout(MainWindow window, double width, double height) {
        window.Width = width; window.Height = height;
        window.ApplyResponsiveLayout(width);
        window.Measure(new Size(width, height)); window.Arrange(new Rect(0, 0, width, height)); window.UpdateLayout();
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();
        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
    }

    private static void Capture(Window window, string name) {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrEmpty(directory)) return;
        Directory.CreateDirectory(directory);
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        frame.Save(Path.Combine(directory, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
