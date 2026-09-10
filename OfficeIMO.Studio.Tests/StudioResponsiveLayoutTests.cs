using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using System.Globalization;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioResponsiveLayoutTests {
    [Theory]
    [InlineData(960)]
    [InlineData(1280)]
    public async Task OcrPanelsReflowWithoutOverlapOrEmptyRows(int width) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var window = new MainWindow();
            try {
                window.Show();
                window.ViewModel.ShowOcrCommand.Execute(null);
                Layout(window, width, 620);
                var workflow = window.GetVisualDescendants().OfType<SearchablePdfOcrView>().Single();
                var source = workflow.FindControl<Border>("SourcePanel")!;
                var information = workflow.FindControl<Control>("InformationPanel")!;
                Assert.False(source.Bounds.Intersects(information.Bounds));
                Assert.True(information.Bounds.Right <= workflow.Bounds.Width + 1);
                if (width == 960) {
                    Assert.InRange(information.Bounds.Top - source.Bounds.Bottom, 0, 1);
                    Assert.Equal(source.Bounds.Left, information.Bounds.Left);
                } else {
                    Assert.Equal(source.Bounds.Top, information.Bounds.Top);
                    Assert.True(information.Bounds.Left >= source.Bounds.Right);
                }
            } finally {
                window.Close();
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData("en", 13)]
    [InlineData("qps-ploc", 17)]
    public async Task DocumentCommandsStayWithinTheWorkspaceWithExpandedLabels(string culture, int fontSize) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            IStudioLocalizer original = StudioLocalization.Current;
            StudioLocalization.Configure(new StudioLocalizer(CultureInfo.GetCultureInfo(culture)));
            var window = new MainWindow { FontSize = fontSize };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory,
                    "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                var workspace = window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!;
                var model = window.ViewModel;
                foreach (var command in new[] { model.ShowViewModeCommand, model.ShowAnnotateModeCommand,
                    model.ShowEditModeCommand, model.ShowPagesModeCommand, model.ShowFormsModeCommand,
                    model.ShowProtectModeCommand }) {
                    command.Execute(null);
                    Layout(window, 960, 620);
                    foreach (var panel in workspace.GetVisualDescendants().OfType<WrapPanel>()
                        .Where(panel => panel.IsEffectivelyVisible)) {
                        foreach (Control child in panel.Children.Where(child => child.IsVisible))
                            Assert.True(child.Bounds.Right <= panel.Bounds.Width + 3,
                                $"{child.GetType().Name} extends beyond toolbar: {child.Bounds}, {panel.Bounds}");
                    }
                    Assert.True(window.ReaderPagesListControl.Bounds.Height >= 180,
                        $"Page viewport is too short in {model.DocumentMode}: {window.ReaderPagesListControl.Bounds}; rows: {string.Join(", ", ((Grid)workspace.Content!).RowDefinitions.Select(row => row.ActualHeight))}");
                }
            } finally {
                window.Close();
                StudioLocalization.Configure(original);
            }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task CompactPanesPreserveCanvasSpaceAndSearchRemainsReachable() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var window = new MainWindow();
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory,
                    "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                Layout(window, 960, 620);
                var workspace = window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!;
                var navigation = workspace.FindControl<Grid>("NavigationPane")!;
                var inspector = workspace.FindControl<Grid>("InspectorPane")!;
                var navigationToggle = workspace.FindControl<ToggleButton>("NavigationToggle")!;
                var inspectorToggle = workspace.FindControl<ToggleButton>("InspectorToggle")!;
                Assert.False(navigation.IsVisible, $"Window {window.Bounds}, workspace {workspace.Bounds}");
                Assert.False(inspector.IsVisible);
                Assert.True(window.ReaderPagesListControl.Bounds.Width >= 800);

                Toggle(navigationToggle);
                Layout(window, 960, 620);
                Assert.True(navigation.IsVisible);
                Assert.False(inspector.IsVisible);
                Assert.True(window.ReaderPagesListControl.Bounds.Width >= 580);

                Toggle(inspectorToggle);
                Layout(window, 960, 620);
                Assert.False(navigation.IsVisible, $"Window {window.Bounds}, workspace {workspace.Bounds}");
                Assert.True(inspector.IsVisible);
                Assert.True(window.ReaderPagesListControl.Bounds.Width >= 520);

                workspace.FocusSearch();
                Layout(window, 960, 620);
                Assert.True(navigation.IsVisible);
                Assert.False(inspector.IsVisible);
                Assert.Equal(2, workspace.FindControl<TabControl>("NavigationTabs")!.SelectedIndex);
                Assert.True(workspace.FindControl<TextBox>("SearchBox")!.IsEffectivelyVisible);
                Assert.True(window.AreFitShortcutsVisible);

                Layout(window, 1600, 900);
                Assert.True(navigation.IsVisible);
                Assert.False(inspector.IsVisible);
                window.ViewModel.ShowAnnotateModeCommand.Execute(null);
                Layout(window, 1600, 900);
                Assert.True(inspector.IsVisible);
                Assert.True(window.ReaderPagesListControl.Bounds.Width >= 900);
            } finally {
                window.Close();
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(960, 620, 13)]
    [InlineData(960, 620, 17)]
    [InlineData(1280, 820, 13)]
    [InlineData(1600, 900, 17)]
    public async Task TaskCardsKeepTheirTextInsideSeparateClickableAreas(int width, int height, int fontSize) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var window = new MainWindow { FontSize = fontSize };
            try {
                window.Show();
                foreach (var command in new[] { window.ViewModel.ShowHomeCommand, window.ViewModel.ShowToolsCommand }) {
                    command.Execute(null);
                    Layout(window, width, height);
                    var panels = window.GetVisualDescendants().OfType<AdaptiveCardPanel>()
                        .Where(panel => panel.IsEffectivelyVisible).ToArray();
                    Assert.NotEmpty(panels);
                    foreach (var panel in panels) {
                        var cards = panel.Children.ToArray();
                        Assert.NotEmpty(cards);
                        foreach (var card in cards) {
                            Assert.True(card is Button || card.GetVisualDescendants().OfType<Button>().Any(),
                                "Each task card must expose a reachable action.");
                            Assert.True(card.Bounds.Right <= panel.Bounds.Width + 3, $"Card {card.Bounds}, panel {panel.Bounds}");
                            Assert.True(card.Bounds.Bottom <= panel.Bounds.Height + 1);
                            foreach (var text in card.GetVisualDescendants().OfType<TextBlock>()) {
                                Point origin = text.TranslatePoint(default, card)!.Value;
                                Assert.True(origin.X >= 0 && origin.X + text.Bounds.Width <= card.Bounds.Width + 1);
                                Assert.True(origin.Y >= 0 && origin.Y + text.Bounds.Height <= card.Bounds.Height + 1, $"Text {text.Text}: {origin}, {text.Bounds}, card {card.Bounds}");
                                foreach (var icon in card.GetVisualDescendants().OfType<PathIcon>()) {
                                    Point iconOrigin = icon.TranslatePoint(default, card)!.Value;
                                    Assert.False(new Rect(origin, text.Bounds.Size).Intersects(new Rect(iconOrigin, icon.Bounds.Size)),
                                        $"Card text overlaps its icon: {text.Text}");
                                }
                            }
                        }
                        for (int i = 0; i < cards.Length; i++)
                            for (int j = i + 1; j < cards.Length; j++)
                                Assert.False(cards[i].Bounds.Intersects(cards[j].Bounds));
                    }
                }
            } finally {
                window.Close();
            }
            return true;
        }, CancellationToken.None);
    }

    private static void Toggle(ToggleButton button) {
        button.IsChecked = button.IsChecked != true;
        button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
    }

    private static void Layout(MainWindow window, double width, double height) {

        window.Width = width;
        window.Height = height;
        window.ApplyResponsiveLayout(width);
        window.Measure(new Size(width, height));
        window.Arrange(new Rect(0, 0, width, height));
        window.UpdateLayout();
        Avalonia.Threading.Dispatcher.UIThread.RunJobs();
        Avalonia.Headless.AvaloniaHeadlessPlatform.ForceRenderTimerTick();
    }
}