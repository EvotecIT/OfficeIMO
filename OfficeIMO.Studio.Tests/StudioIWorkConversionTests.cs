using System.IO.Compression;
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
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioIWorkConversionTests {
    [Theory]
    [InlineData("simple.pages", "pages-docx", false)]
    [InlineData("simple.pages", "pages-docx", true)]
    [InlineData("simple.numbers", "numbers-xlsx", false)]
    [InlineData("simple.numbers", "numbers-xlsx", true)]
    [InlineData("tabledeck.key", "keynote-pptx", false)]
    [InlineData("tabledeck.key", "keynote-pptx", true)]
    public async Task Studio_default_registry_intakes_Apple_formats_and_honors_conversion_acceptance(string fixture, string route, bool directory) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, fixture);
            if (directory) ZipFile.ExtractToDirectory(Fixture(fixture), source);
            else File.Copy(Fixture(fixture), source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var queue = model.ConversionWorkbench;
            Assert.True(queue.AddDroppedPaths([source]));
            ConversionJobViewModel job = Assert.Single(queue.Jobs);
            Assert.Equal(route, job.Route.Route.Id);
            Assert.True(job.SupportsIWorkOptions);
            Assert.False(job.AllowIncompleteVisualPreview);
            Assert.False(job.AllowPartialEditableReconstruction);
            await queue.RunQueueCommand.ExecuteAsync(null);
            // These native styles contain paragraph layout that requires explicit partial acceptance.
            Assert.Equal(ConversionJobState.Failed, job.State);
            Assert.False(job.HasOutput);
            job.AllowPartialEditableReconstruction = true;
            await queue.RetryFailedCommand.ExecuteAsync(null);
            if (route == "keynote-pptx") {
                Assert.Contains(job.ConversionEvidence!.FidelityDiagnostics,
                    diagnostic => diagnostic.Code == "IWORK_KEYNOTE_PARAGRAPH_PAGINATION_OMITTED");
            }
            Assert.Equal(ConversionJobState.Completed, job.State);
            Assert.True(job.HasOutput, job.Summary);
            if (route == "keynote-pptx") {
                using var saved = OfficeIMO.PowerPoint.PowerPointPresentation.Load(job.OutputPath!);
                var table = Assert.Single(saved.Slides.SelectMany(slide => slide.Tables));
                Assert.Equal("Widget", table.GetCell(1, 0).Text);
                for (int row = 0; row < 3; row++)
                    for (int column = 0; column < 3; column++) Assert.True(table.GetCell(row, column).NoFill);
            }
            Assert.True(job.HasConversionEvidence);
            Assert.Equal(64, job.SourceFingerprint.Length);
            Assert.Contains(job.Diagnostics, item => item.Code == "SourceSnapshot" && item.Details.GetValueOrDefault("snapshotKind") == (directory ? "DirectoryPackage" : "FileBytes"));
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
                acceptance.BringIntoView();
                await Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), DispatcherPriority.Background);
                Assert.True(acceptance.IsEffectivelyVisible && acceptance.IsEnabled);
                Point point = acceptance.TranslatePoint(new Point(10, acceptance.Bounds.Height / 2), window)!.Value;
                Assert.InRange(point.X, 0, width); Assert.InRange(point.Y, 0, height);
                var hit = window.InputHitTest(point) as Visual;
                Assert.True(ReferenceEquals(hit, acceptance)
                    || hit?.GetVisualAncestors().Contains(acceptance) == true,
                    "The acceptance checkbox must be reachable by pointer inside the scroll viewport.");
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
                var facts = job.ConversionEvidence.Facts;
                Assert.Equal("0", facts["reconstructedSourceUnitCount"]);
                Assert.Equal(facts["sourceUnitCount"], facts["unassessedSourceUnitCount"]);
                Assert.Contains($"{facts["sourceUnitCount"]} identified source units", job.ConversionEvidenceSummary);
                Assert.Contains($"{facts["unassessedSourceUnitCount"]} unassessed", job.ConversionEvidenceSummary);
                await queue.PreviewOutputCommand.ExecuteAsync(null);
                Assert.NotEmpty(queue.OutputPreviewPages);
                Assert.Contains("saved artifact", queue.OutputPreviewStatus, StringComparison.OrdinalIgnoreCase);
                view.FindControl<StackPanel>("ConversionEvidencePanel")!.BringIntoView(); window.UpdateLayout();
                Capture(window, "apple-result-" + width);
                return true;
            } finally { window.Close(); }
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(840, 600)]
    [InlineData(1280, 800)]
    public async Task Numbers_source_formula_assessments_are_visible_after_conversion(int width, int height) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = width > 1000 ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "formulas.numbers");
            File.Copy(Fixture("formulas.numbers"), source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var queue = model.ConversionWorkbench;
            queue.AddDroppedPaths([source]);
            var job = Assert.Single(queue.Jobs);
            job.AllowPartialEditableReconstruction = true;
            var view = new ConversionWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show(); window.UpdateLayout();
                await queue.RunQueueCommand.ExecuteAsync(null);
                Assert.Equal(ConversionJobState.Completed, job.State);
                Assert.Equal("28", job.ConversionEvidence!.Facts["sourceFormulaCellCount"]);
                Assert.Equal("28", job.ConversionEvidence.Facts["sourceCompleteFormulaExpressionCount"]);
                Assert.Contains("Source formulas: 28 complete", job.ConversionEvidenceSummary);
                Assert.Contains("Recovered caches:", job.ConversionEvidenceSummary);
                await queue.PreviewOutputCommand.ExecuteAsync(null);
                Assert.NotEmpty(queue.OutputPreviewPages);
                view.FindControl<StackPanel>("ConversionEvidencePanel")!.BringIntoView(); window.UpdateLayout();
                Capture(window, "apple-formulas-" + width);
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

    [Theory]
    [InlineData(840, 600)]
    [InlineData(1280, 800)]
    public async Task Unassessed_formula_evidence_is_visible_in_the_rendered_summary(int width, int height) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = width > 1000 ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "unassessed.numbers");
            File.Copy(Fixture("simple.numbers"), source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var queue = model.ConversionWorkbench;
            queue.AddDroppedPaths([source]);
            var job = Assert.Single(queue.Jobs);
            job.AllowPartialEditableReconstruction = true;
            await queue.RunQueueCommand.ExecuteAsync(null);
            Assert.Equal(ConversionJobState.Completed, job.State);
            // Supply the retained workflow-fact boundary for a single undecoded declaration.
            // Core package tests establish this source state; this test exercises its UI consumer.
            var evidence = job.ConversionEvidence!;
            var facts = new Dictionary<string, string>(evidence.Facts) {
                ["sourceFormulaCellCount"] = "1",
                ["sourceCompleteFormulaExpressionCount"] = "0",
                ["sourceIncompleteFormulaExpressionCount"] = "0",
                ["sourceUnassessedFormulaExpressionCount"] = "1",
                ["sourceCompleteFormulaCacheCount"] = "0",
                ["sourcePartialFormulaCacheCount"] = "0",
                ["sourceApproximateFormulaCacheCount"] = "0",
                ["sourceMissingFormulaCacheCount"] = "0",
                ["sourceUnassessedFormulaCacheCount"] = "1"
            };
            job.ConversionEvidence = new OfficeWorkflowConversionEvidence(evidence, facts);
            var view = new ConversionWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show(); window.UpdateLayout();
                var panel = view.FindControl<StackPanel>("ConversionEvidencePanel")!;
                panel.BringIntoView(); window.UpdateLayout();
                Assert.True(panel.IsEffectivelyVisible);
                string summary = job.ConversionEvidenceSummary;
                Assert.Contains("Source formulas: 0 complete · 0 incomplete · 1 unassessed", summary);
                Assert.Contains("Recovered caches: 0 complete · 0 partial · 0 approximate · 0 missing · 1 unassessed", summary);
                Assert.Contains(panel.GetVisualDescendants().OfType<TextBlock>(), block => block.Text == summary);
                Capture(window, "apple-unassessed-formulas-" + width);
                return true;
            } finally { window.Close(); }
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
