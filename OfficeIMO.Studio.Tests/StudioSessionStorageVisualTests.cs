using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioSessionStorageVisualTests {
    [Theory]
    [InlineData(960, 620, false, false)]
    [InlineData(1280, 820, true, false)]
    [InlineData(960, 620, false, true)]
    [InlineData(1280, 820, true, true)]
    public async Task SessionStorageFailureStaysVisibleUntilRetrySucceeds(int width, int height, bool dark, bool clearSession) {
        if (!OperatingSystem.IsWindows()) return; // Windows file sharing denies replacement while held open.
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string source = Path.Combine(services.Paths.Root, "session-storage.pdf");
            PdfDocument.Create(builder => builder.Page(page => page.Size(600, 800)
                .Content(content => content.Text("The open document remains available during a storage failure.")))).Save(source);
            byte[] original = File.ReadAllBytes(source);
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(source);
                var session = window.ViewModel.Session!;
                session.Flush();
                byte[] previousSession = File.ReadAllBytes(services.Paths.SessionPath);
                Button retry;
                using (var held = new FileStream(services.Paths.SessionPath, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                    if (clearSession) window.ViewModel.Settings.RememberSession = false;
                    else session.Flush();
                    Assert.True(session.HasStorageError);
                    Layout(window, width, height);
                    retry = window.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, session.RetryStorageCommand));
                    Assert.True(retry.IsEffectivelyVisible);
                    Point topLeft = retry.TranslatePoint(default, window)!.Value;
                    Assert.InRange(topLeft.X, 0, width - retry.Bounds.Width);
                    Assert.InRange(topLeft.Y, 0, height - retry.Bounds.Height);
                    retry.Command!.Execute(null);
                    Assert.True(session.HasStorageError);
                    Assert.Equal(previousSession, File.ReadAllBytes(services.Paths.SessionPath));
                    await window.ViewModel.Pages[0].EnsureRenderedAsync();
                    Capture(window, width, dark, clearSession ? "clear-failed" : "failed");
                }
                // A successful storage retry must not dismiss an unrelated restore/recovery error.
                session.Error = "A separate recovery operation failed.";
                retry.Command!.Execute(null);
                Assert.False(session.HasStorageError);
                Assert.True(session.HasError);
                Layout(window, width, height);
                Assert.False(retry.IsEffectivelyVisible);
                Assert.Equal(!clearSession, File.Exists(services.Paths.SessionPath));
                Assert.Equal(original, File.ReadAllBytes(source));
                Capture(window, width, dark, clearSession ? "clear-retried" : "retried");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void Layout(Window window, int width, int height) {
        window.Measure(new Size(width, height));
        window.Arrange(new Rect(0, 0, width, height));
        window.UpdateLayout();
    }

    private static void Capture(Window window, int width, bool dark, string state) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, $"session-storage-{width}-{(dark ? "dark" : "light")}-{state}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
