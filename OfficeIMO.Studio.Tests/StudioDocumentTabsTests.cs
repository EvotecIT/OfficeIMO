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
