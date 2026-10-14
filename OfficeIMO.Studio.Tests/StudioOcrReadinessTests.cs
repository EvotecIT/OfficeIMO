using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Settings;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioOcrReadinessTests {
    [Fact]
    public async Task Missing_engine_is_explained_on_the_ocr_screen_and_leads_to_its_setup() {
        if (StudioOcrProvider.UnavailableReason is not null) return;
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var services = ((App)Application.Current!).Services;
            // A saved engine location that no longer exists is the same as no engine at all.
            services.Preferences.Update(current => current with {
                OcrExecutablePath = Path.Combine(services.Paths.Root, "missing", OperatingSystem.IsWindows() ? "tesseract.exe" : "tesseract")
            });
            var window = new MainWindow(services) { Width = 1280, Height = 820 };
            try {
                window.Show();
                MainWindowViewModel document = window.ViewModel;
                Assert.True(document.OcrReadiness.HasNotice);
                Assert.False(document.OcrReadiness.IsReady);
                Assert.False(document.CanDetectPageOrientation);
                Assert.False(document.OcrSession.IsRecognitionAvailable);

                document.ShowOcrCommand.Execute(null);
                Assert.True(document.OcrReadiness.HasNotice);
                Assert.True(document.OcrReadiness.CanSetUp);
                Assert.Equal(services.Localizer.Get("Ocr.Notice.EngineMissing"), document.OcrReadiness.Notice);
                Assert.Same(document.OcrReadiness, document.OcrSession.Readiness);

                document.OcrReadiness.SetUpCommand.Execute(null);
                Assert.True(window.ViewModel.IsSettingsMode);
                for (int attempt = 0; attempt < 20 && document.Settings.RequestedSection is not null; attempt++) Layout(window);
                Assert.Null(document.Settings.RequestedSection);
                var list = window.GetVisualDescendants().OfType<ListBox>().Single(box => box.Name == "SectionList");
                Assert.Equal(StudioSettingsViewModel.TextRecognitionSection, (list.SelectedItem as ListBoxItem)?.Tag);

                // Choosing automatic discovery again rechecks; the notice follows the saved choice.
                services.Preferences.Update(current => current with { OcrExecutablePath = null });
                Assert.Equal(!Ocr.Tesseract.TesseractRuntime.TryDiscover(null, out _), document.OcrReadiness.HasNotice);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public void Readiness_reports_nothing_when_the_engine_is_found() {
        if (StudioOcrProvider.UnavailableReason is not null) return;
        using var readiness = new OcrReadinessViewModel(null, TestAppBuilder.CreateTestServices().Localizer, _ => true);
        readiness.Refresh();
        Assert.False(readiness.HasNotice);
        Assert.False(readiness.CanSetUp);
        Assert.True(readiness.IsReady);
    }

    private static void Layout(Window window) {
        Dispatcher.UIThread.RunJobs();
        window.Measure(new Size(1280, 820));
        window.Arrange(new Rect(0, 0, 1280, 820));
        window.UpdateLayout();
        Dispatcher.UIThread.RunJobs();
    }
}
