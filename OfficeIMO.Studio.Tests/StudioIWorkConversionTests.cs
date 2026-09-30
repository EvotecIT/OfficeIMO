using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.IWork;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioIWorkConversionTests {
    [Theory]
    [InlineData("simple.pages", "pages-docx")]
    [InlineData("simple.numbers", "numbers-xlsx")]
    [InlineData("tabledeck.key", "keynote-pptx")]
    public async Task Studio_default_registry_intakes_and_converts_all_Apple_formats(string fixture, string route) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, fixture);
            File.Copy(Fixture(fixture), source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var queue = model.ConversionWorkbench;
            Assert.True(queue.AddDroppedPaths([source]));
            ConversionJobViewModel job = Assert.Single(queue.Jobs);
            Assert.Equal(route, job.Route.Route.Id);
            Assert.True(job.SupportsIWorkOptions);
            Assert.False(job.AllowIncompleteVisualPreview);
            Assert.False(job.AllowPartialEditableReconstruction);
            await queue.RunQueueCommand.ExecuteAsync(null);
            Assert.Equal(ConversionJobState.Completed, job.State);
            Assert.True(job.HasOutput, job.Summary);
            Assert.True(job.HasConversionEvidence);
            Assert.Equal(64, job.SourceFingerprint.Length);
            Assert.Contains(job.Diagnostics, diagnostic => diagnostic.Details.GetValueOrDefault("lossKind") == "Unassessed");
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(840, 600)]
    [InlineData(1280, 800)]
    public async Task Preview_coverage_requires_visible_acceptance_and_retains_saved_output_evidence(int width, int height) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = width > 1000 ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "review.pages");
            File.Copy(Fixture("simple.pages"), source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var queue = model.ConversionWorkbench;
            queue.AddDroppedPaths([source]);
            var job = Assert.Single(queue.Jobs);
            job.IWorkMode = IWorkConversionMode.VisualOnly;
            var view = new ConversionWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show();
                window.UpdateLayout();
                Assert.True(view.FindControl<StackPanel>("IWorkAcceptancePanel")!.IsEffectivelyVisible);
                Assert.Equal(IWorkConversionMode.VisualOnly, view.FindControl<ComboBox>("IWorkModeChoice")!.SelectedItem);
                await queue.RunQueueCommand.ExecuteAsync(null);
                Assert.Equal(ConversionJobState.Failed, job.State);
                Assert.False(job.HasOutput);
                Assert.Contains("complete", job.Summary!, StringComparison.OrdinalIgnoreCase);
                var acceptance = view.FindControl<CheckBox>("IWorkPreviewChoice")!;
                var detailsScroll = acceptance.GetVisualAncestors().OfType<ScrollViewer>().First();
                Point scrollPoint = detailsScroll.TranslatePoint(new Point(detailsScroll.Bounds.Width / 2, detailsScroll.Bounds.Height / 2), window)!.Value;
                window.MouseWheel(scrollPoint, new Vector(0, -4));
                await Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), DispatcherPriority.Background);
                Assert.True(acceptance.IsEffectivelyVisible && acceptance.IsEnabled);
                Point point = acceptance.TranslatePoint(new Point(10, acceptance.Bounds.Height / 2), window)!.Value;
                Assert.InRange(point.X, 0, width); Assert.InRange(point.Y, 0, height);
                Capture(window, "apple-acceptance-" + width);
                window.MouseDown(point, MouseButton.Left); window.MouseUp(point, MouseButton.Left);
                Assert.True(job.AllowIncompleteVisualPreview);
                await queue.RetryFailedCommand.ExecuteAsync(null);
                Assert.Equal(ConversionJobState.Completed, job.State);
                Assert.Equal("VisualFallback", job.ConversionEvidence!.Facts["projectionKind"]);
                Assert.Contains(job.ConversionEvidence.FidelityDiagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
                Assert.True(view.FindControl<StackPanel>("ConversionEvidencePanel")!.IsEffectivelyVisible);
                Assert.Equal(job.SourceFingerprint, view.FindControl<TextBox>("SourceFingerprintText")!.Text);
                Assert.Contains("incomplete", job.ConversionEvidenceSummary, StringComparison.OrdinalIgnoreCase);
                await queue.PreviewOutputCommand.ExecuteAsync(null);
                Assert.NotEmpty(queue.OutputPreviewPages);
                Assert.Contains("saved artifact", queue.OutputPreviewStatus, StringComparison.OrdinalIgnoreCase);
                view.FindControl<StackPanel>("ConversionEvidencePanel")!.BringIntoView(); window.UpdateLayout();
                Capture(window, "apple-result-" + width);
                return true;
            } finally { window.Close(); }
        }, CancellationToken.None);
    }

    [Fact]
    public async Task Startup_Apple_document_enters_conversion_instead_of_a_PDF_tab() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "startup.numbers");
            File.Copy(Fixture("simple.numbers"), source);
            var window = new MainWindow(services);
            try {
                window.OpenInitialDocument([source]); window.Show();
                var deadline = DateTime.UtcNow.AddSeconds(5);
                while (window.IsStartingUp && DateTime.UtcNow < deadline) await Task.Delay(10);
                Assert.False(window.IsStartingUp);
                Assert.True(window.ViewModel.IsConversionMode);
                Assert.Equal("numbers-xlsx", Assert.Single(window.ViewModel.ConversionWorkbench.Jobs).Route.Route.Id);
                Assert.Empty(window.ViewModel.Pages);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "IWork", name);
    private static void Capture(Window window, string name) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        frame.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
