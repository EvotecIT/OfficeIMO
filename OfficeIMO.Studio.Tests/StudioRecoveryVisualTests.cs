using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Settings;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioRecoveryVisualTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(960, 620, true)]
    [InlineData(1280, 820, false)]
    [InlineData(1280, 820, true)]
    public async Task RecoveryConfirmationAndResultStayVisibleAtSupportedSizes(int width, int height, bool dark) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string source = Path.Combine(services.Paths.Root, "visual-source.pdf");
            byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
            File.WriteAllBytes(source, bytes);
            string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint(bytes);
            await services.Recovery.WriteAsync(source, fingerprint, bytes, 1, CancellationToken.None);
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                window.ViewModel.ShowSettingsCommand.Execute(null);
                var settings = window.ViewModel.Settings;
                settings.RequestRecoveryClearCommand.Execute(null);
                Layout(window, width, height);
                SettingsView view = window.GetVisualDescendants().OfType<SettingsView>().Single();
                var scroll = (ScrollViewer)view.Content!;
                scroll.Offset = new Vector(0, scroll.Extent.Height);
                Layout(window, width, height);
                var clear = view.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, settings.ClearRecoveryCommand));
                var cancel = view.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, settings.CancelRecoveryClearCommand));
                Rect clearBounds = VisibleBounds(clear, view);
                Rect cancelBounds = VisibleBounds(cancel, view);
                Assert.False(clearBounds.Intersects(cancelBounds));
                Capture(window, width, dark, "confirm");

                settings.CancelRecoveryClearCommand.Execute(null);
                Assert.False(settings.ConfirmRecoveryClear);
                Assert.NotNull(services.Recovery.Find(source, fingerprint));
                settings.RequestRecoveryClearCommand.Execute(null);
                await settings.ClearRecoveryCommand.ExecuteAsync(null);
                Layout(window, width, height);
                Assert.True(settings.HasRecoveryStatus);
                Assert.Null(services.Recovery.Find(source, fingerprint));
                Assert.Equal(bytes, File.ReadAllBytes(source));
                TextBlock status = view.GetVisualDescendants().OfType<TextBlock>().Single(text => text.Text == settings.RecoveryStatus);
                VisibleBounds(status, view);
                Capture(window, width, dark, "cleared");
            } finally {
                window.Close();
            }
            return true;
        }, CancellationToken.None);
    }

    private static void Layout(Window window, int width, int height) {
        Dispatcher.UIThread.RunJobs();
        window.Measure(new Size(width, height));
        window.Arrange(new Rect(0, 0, width, height));
        window.UpdateLayout();
    }

    private static Rect VisibleBounds(Control control, Control ancestor) {
        Point position = control.TranslatePoint(default, ancestor)!.Value;
        var bounds = new Rect(position, control.Bounds.Size);
        Assert.True(control.IsEffectivelyVisible);
        Assert.True(bounds.Top >= -1 && bounds.Bottom <= ancestor.Bounds.Height + 1);
        Assert.True(bounds.Left >= -1 && bounds.Right <= ancestor.Bounds.Width + 1);
        return bounds;
    }

    private static void Capture(Window window, int width, bool dark, string state) {
        using var bitmap = window.CaptureRenderedFrame();
        Assert.NotNull(bitmap);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        bitmap.Save(Path.Combine(output, $"recovery-{(dark ? "dark" : "light")}-{width}-{state}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
