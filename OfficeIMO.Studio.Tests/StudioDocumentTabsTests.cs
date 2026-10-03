using Avalonia;
using Avalonia.Automation.Peers;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Input.Raw;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioDocumentTabsTests {
    [Fact]
    public async Task DocumentTabsExposeSelectionAndKeepKeyboardNavigationVisible() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-tab-controls-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var window = new MainWindow();
                try {
                    window.Show();
                    for (int index = 0; index < 8; index++) {
                        string path = Path.Combine(root, $"Document {index + 1}.pdf");
                        File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"), path);
                        await window.TabHost.OpenDocumentAsync(path);
                    }
                    Layout(window, 900);
                    var tabs = window.FindControl<TabControl>("DocumentTabs")!;
                    Assert.Equal(AutomationControlType.Tab, ControlAutomationPeer.CreatePeerForElement(tabs)!.GetAutomationControlType());
                    var last = Assert.IsType<TabItem>(tabs.ContainerFromIndex(7));
                    Assert.Equal(AutomationControlType.TabItem, ControlAutomationPeer.CreatePeerForElement(last)!.GetAutomationControlType());
                    Assert.Equal("Document 8.pdf", ControlAutomationPeer.CreatePeerForElement(last)!.GetName());
                    AssertVisibleInStrip(tabs, last);

                    last.Focus();
                    window.KeyPress(Key.Home, RawInputModifiers.None, PhysicalKey.None, null);
                    Layout(window, 900);
                    Assert.Same(window.TabHost.Tabs[0], window.TabHost.SelectedTab);
                    AssertVisibleInStrip(tabs, tabs.ContainerFromIndex(0)!);
                    window.KeyPress(Key.Right, RawInputModifiers.None, PhysicalKey.None, null);
                    Assert.Same(window.TabHost.Tabs[1], window.TabHost.SelectedTab);
                    window.KeyPress(Key.End, RawInputModifiers.None, PhysicalKey.None, null);
                    Layout(window, 900);
                    Assert.Same(window.TabHost.Tabs[7], window.TabHost.SelectedTab);
                    AssertVisibleInStrip(tabs, last);

                    // A menu/programmatic selection must reveal its tab without moving keyboard focus.
                    var search = window.FindControl<Button>("CommandSearchButton")!;
                    search.Focus();
                    window.TabHost.SelectedTab = window.TabHost.Tabs[0];
                    Layout(window, 900);
                    Assert.Same(search, window.FocusManager!.GetFocusedElement());
                    AssertVisibleInStrip(tabs, tabs.ContainerFromIndex(0)!);
                    var close = tabs.ContainerFromIndex(0)!.GetVisualDescendants().OfType<Button>().Single();
                    Assert.Equal("Close Document 1.pdf", ControlAutomationPeer.CreatePeerForElement(close)!.GetName());

                    Layout(window, 390);
                    var picker = window.FindControl<ComboBox>("CompactDocumentPicker")!;
                    Assert.True(picker.IsEffectivelyVisible);
                    picker.SelectedIndex = 7;
                    Assert.Same(window.TabHost.Tabs[7], window.TabHost.SelectedTab);
                    Layout(window, 700);
                    AssertVisibleInStrip(tabs, tabs.ContainerFromIndex(7)!);

                    tabs.ContainerFromIndex(7)!.Focus();
                    var closed = window.TabHost.SelectedTab;
                    window.KeyPress(Key.W, OperatingSystem.IsMacOS() ? RawInputModifiers.Meta : RawInputModifiers.Control, PhysicalKey.None, null);
                    Layout(window, 1280);
                    Assert.DoesNotContain(closed, window.TabHost.Tabs);
                    Assert.Same(tabs.ContainerFromIndex(tabs.SelectedIndex), window.FocusManager!.GetFocusedElement());
                } finally { window.Close(); }
                return true;
            }, CancellationToken.None);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ReorderMenuAndCompactClosePreserveLiveDocuments() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-tab-reorder-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var window = new MainWindow();
                try {
                    window.Show();
                    for (int index = 0; index < 3; index++) {
                        string path = Path.Combine(root, $"Document {index + 1}.pdf");
                        File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"), path);
                        await window.TabHost.OpenDocumentAsync(path);
                    }
                    Layout(window, 1280);
                    var tabs = window.FindControl<TabControl>("DocumentTabs")!;
                    var first = window.TabHost.Tabs[0];
                    var second = window.TabHost.Tabs[1];
                    var third = window.TabHost.Tabs[2];
                    string? thirdPath = third.SourcePath;
                    Point start = tabs.ContainerFromIndex(0)!.TranslatePoint(new Point(80, 16), window)!.Value;
                    var last = tabs.ContainerFromIndex(2)!;
                    Point end = last.TranslatePoint(new Point(last.Bounds.Width - 8, 16), window)!.Value;
                    window.MouseDown(start, MouseButton.Left);
                    window.MouseMove(end);
                    window.MouseUp(end, MouseButton.Left);
                    Layout(window, 1280);
                    Assert.Equal(new[] { second, third, first }, window.TabHost.Tabs);
                    Assert.Same(first, window.TabHost.SelectedTab);
                    Assert.Same(first.Document, window.ViewModel);
                    Assert.True(first.Document.HasDocument);

                    // Reordering is available without a pointer, and preserves the selected session.
                    tabs.ContainerFromIndex(2)!.Focus();
                    window.KeyPress(Key.Left, RawInputModifiers.Alt | RawInputModifiers.Shift, PhysicalKey.None, null);
                    Layout(window, 1280);
                    Assert.Equal(new[] { second, first, third }, window.TabHost.Tabs);
                    Assert.Same(first.Document, window.ViewModel);
                    var contextTarget = Assert.IsType<TabItem>(tabs.ContainerFromIndex(1));
                    contextTarget.RaiseEvent(new ContextRequestedEventArgs());
                    var context = Assert.IsType<MenuFlyout>(contextTarget.ContextFlyout);
                    Assert.True(context.IsOpen);
                    Assert.Equal(3, context.Items.Count);
                    context.Hide();

                    // Releasing outside the strip cancels a drag instead of moving or closing a document.
                    start = tabs.ContainerFromIndex(1)!.TranslatePoint(new Point(80, 16), window)!.Value;
                    end = start + new Vector(100, 100);
                    window.MouseDown(start, MouseButton.Left);
                    window.MouseMove(end, RawInputModifiers.LeftMouseButton);
                    window.MouseUp(end, MouseButton.Left);
                    Layout(window, 1280);
                    Assert.Equal(new[] { second, first, third }, window.TabHost.Tabs);

                    var button = window.FindControl<Button>("DocumentListButton")!;
                    var menu = Assert.IsType<MenuFlyout>(button.Flyout);
                    Point menuPoint = button.TranslatePoint(new Point(15, 15), window)!.Value;
                    window.MouseDown(menuPoint, MouseButton.Left);
                    window.MouseUp(menuPoint, MouseButton.Left);
                    Layout(window, 1280);
                    Assert.True(menu.IsOpen);
                    var presenter = Assert.IsAssignableFrom<ItemsControl>(menu.Popup.Child);
                    Assert.Equal(menu.Items.Count, presenter.ItemCount);
                    Assert.True(presenter.Bounds.Width > 100 && presenter.Bounds.Height > 100);
                    var entries = menu.Items.OfType<MenuItem>().Take(3).ToArray();
                    Assert.Equal(new[] { second.Title, first.Title, third.Title }, entries.Select(item => item.Header));
                    Assert.True(entries[1].IsChecked);
                    entries[2].RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(MenuItem.ClickEvent));
                    menu.Hide();
                    Layout(window, 1280);
                    Assert.Same(third, window.TabHost.SelectedTab);
                    Capture(window, "workspace-reordered-wide");

                    Layout(window, 390);
                    var close = window.FindControl<Button>("CompactCloseDocumentButton")!;
                    Assert.True(close.IsEffectivelyVisible);
                    Assert.True(close.Bounds.Width >= 44);
                    Capture(window, "workspace-compact-close");
                    close.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
                    Layout(window, 390);
                    Assert.DoesNotContain(third, window.TabHost.Tabs);
                    Assert.Equal(2, window.TabHost.Tabs.Count);
                    await window.TabHost.ReopenClosedTabAsync();
                    Assert.Equal(3, window.TabHost.Tabs.Count);
                    Assert.Equal(thirdPath, window.TabHost.SelectedTab!.SourcePath);
                    while (window.TabHost.HasTabs) {
                        close.Focus();
                        close.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
                        Layout(window, 390);
                    }
                    Assert.Same(window.FindControl<Button>("OpenDocumentTabButton"), window.FocusManager!.GetFocusedElement());
                } finally { window.Close(); }
                return true;
            }, CancellationToken.None);
        } finally { Directory.Delete(root, true); }
    }

    private static void Capture(Window window, string name) {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrEmpty(directory)) return;
        Directory.CreateDirectory(directory);
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        frame.Save(Path.Combine(directory, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }

    [Fact]
    public async Task ReaderSelectionsExposeReadableValuesInsteadOfModelNames() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var window = new MainWindow();
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                Layout(window, 1280);
                var page = window.ViewModel.ReaderPages[0];
                Assert.Equal(page.PageLabel, page.ToString());
                var pages = window.GetVisualDescendants().OfType<ListBox>().Single(control => control.Name == "PagesList");
                Assert.Equal(page.PageLabel, ControlAutomationPeer.CreatePeerForElement(pages.ContainerFromIndex(0)!)!.GetName());
                Assert.All(window.ViewModel.ReaderLayoutChoices, choice => Assert.Equal(choice.Label, choice.ToString()));
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void AssertVisibleInStrip(TabControl tabs, Control item) {
        Point origin = item.TranslatePoint(default, tabs)!.Value;
        Assert.True(origin.X >= -1 && origin.X + item.Bounds.Width <= tabs.Bounds.Width + 1,
            $"Tab {origin} {item.Bounds} outside strip {tabs.Bounds}");
    }

    private static void Layout(MainWindow window, double width) {
        window.Width = width; window.Height = 820;
        window.ApplyResponsiveLayout(width);
        window.Measure(new Size(width, 820)); window.Arrange(new Rect(0, 0, width, 820)); window.UpdateLayout();
        Dispatcher.UIThread.RunJobs();
        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
    }
}
