using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Settings;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioSettingsNavigationTests {
    [Fact]
    public async Task Choosing_a_section_scrolls_to_it_and_scrolling_moves_the_selection() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var services = ((App)Application.Current!).Services;
            var window = new MainWindow(services) { Width = 1280, Height = 720 };
            try {
                window.Show();
                window.ViewModel.ShowSettingsCommand.Execute(null);
                Layout(window);
                SettingsView view = window.GetVisualDescendants().OfType<SettingsView>().Single();
                var list = view.GetVisualDescendants().OfType<ListBox>().Single(box => box.Name == "SectionList");
                var scroller = view.GetVisualDescendants().OfType<ScrollViewer>().Single(viewer => viewer.Name == "SectionScroller");
                var sections = list.Items.OfType<ListBoxItem>()
                    .Select(item => view.GetVisualDescendants().OfType<Control>().Single(control => control.Name == (string)item.Tag!))
                    .ToArray();
                Assert.Equal(6, sections.Length);
                Assert.Equal(0, list.SelectedIndex);

                // Choosing a section brings its heading to the top of the page.
                list.SelectedIndex = 2;
                Layout(window);
                double top = sections[2].TranslatePoint(default, scroller)!.Value.Y;
                Assert.InRange(top, 0, 48);
                Assert.Equal(2, list.SelectedIndex);

                // Reading on moves the highlight with the page, and the end of the page selects the last section.
                scroller.Offset = new Vector(0, scroller.Extent.Height);
                Layout(window);
                Assert.Equal(sections.Length - 1, list.SelectedIndex);
                scroller.Offset = default;
                Layout(window);
                Assert.Equal(0, list.SelectedIndex);

                // A narrow window gives the page the full width.
                window.Width = 600;
                Layout(window, 600);
                Assert.False(list.IsEffectivelyVisible);
                Assert.True(sections.All(section => section.IsEffectivelyVisible));
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void Layout(Window window, int width = 1280) {
        Dispatcher.UIThread.RunJobs();
        window.Measure(new Size(width, 720));
        window.Arrange(new Rect(0, 0, width, 720));
        window.UpdateLayout();
        Dispatcher.UIThread.RunJobs();
    }
}
