using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Assistant;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class AssistantLayoutTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(960, 620, true)]
    [InlineData(1600, 900, false)]
    public async Task ConnectionSetupKeepsQuestionAndCancellationAreaReachable(int width, int height, bool light) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(() => {
            var services = TestAppBuilder.CreateTestServices();
            var window = new MainWindow(services) { Width = width, Height = height,
                RequestedThemeVariant = light ? ThemeVariant.Light : ThemeVariant.Dark };
            try {
                window.Show();
                window.ViewModel.ToggleAssistantCommand.Execute(null);
                services.AiConnections.ProviderIndex = 2;
                window.ApplyResponsiveLayout(width);
                window.Measure(new Size(width, height)); window.Arrange(new Rect(0, 0, width, height));
                window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                using var frame = window.CaptureRenderedFrame();
                Assert.NotNull(frame);
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(output)) {
                    Directory.CreateDirectory(output);
                    frame.Save(Path.Combine(output, $"assistant-{width}x{height}-{(light ? "light" : "dark")}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
                var view = Assert.Single(window.GetVisualDescendants().OfType<AssistantView>());
                var ask = Assert.Single(view.GetVisualDescendants().OfType<Button>(), button =>
                    ReferenceEquals(button.Command, window.ViewModel.Assistant.AskCommand));
                var question = Assert.Single(view.GetVisualDescendants().OfType<TextBox>(), box => box.AcceptsReturn);
                foreach (Control control in new Control[] { ask, question }) {
                    Point position = control.TranslatePoint(default, window)!.Value;
                    Assert.True(position.X >= 0 && position.X + control.Bounds.Width <= width + 1);
                    Assert.True(position.Y >= 0 && position.Y + control.Bounds.Height <= height + 1, $"{control.GetType().Name}: position {position}, bounds {control.Bounds}, window {window.Bounds}");
                    Assert.True(control.Bounds.Height >= 30);
                }
                Assert.Equal(width >= 1500 ? SplitViewDisplayMode.Inline : SplitViewDisplayMode.Overlay,
                    window.FindControl<SplitView>("AssistantHost")!.DisplayMode);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
