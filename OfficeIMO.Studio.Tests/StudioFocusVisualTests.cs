using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioFocusVisualTests {
    [Fact]
    public async Task Keyboard_focus_draws_the_shared_accent_ring() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var services = TestAppBuilder.CreateTestServices();
            var window = new MainWindow(services) { Width = 1280, Height = 800 };
            try {
                window.Show();
                window.UpdateLayout();
                Button open = window.GetVisualDescendants().OfType<Button>()
                    .First(button => button.IsEffectivelyVisible && button.Classes.Contains("primary"));
                open.Focus(NavigationMethod.Tab);
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();
                AvaloniaHeadlessPlatform.ForceRenderTimerTick();

                AdornerLayer layer = Assert.IsType<AdornerLayer>(AdornerLayer.GetAdornerLayer(open));
                Border ring = layer.Children.OfType<Control>().SelectMany(child => child.GetSelfAndVisualDescendants().OfType<Border>())
                    .First(border => border.BorderThickness == new Thickness(2));
                Assert.True(window.TryFindResource("StudioAccentBrush", window.ActualThemeVariant, out object? accent));
                Assert.Equal(((ISolidColorBrush)accent!).Color, Assert.IsAssignableFrom<ISolidColorBrush>(ring.BorderBrush).Color);
                Assert.Equal(new CornerRadius(9), ring.CornerRadius);
                window.UpdateLayout();
                Assert.True(ring.Bounds.Width >= open.Bounds.Width, $"ring {ring.Bounds} button {open.Bounds} parent {(ring.GetVisualParent() as Control)?.Bounds}");

                window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                using var frame = window.CaptureRenderedFrame();
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(output)) {
                    Directory.CreateDirectory(output);
                    frame!.Save(Path.Combine(output, "focus-ring-home.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
