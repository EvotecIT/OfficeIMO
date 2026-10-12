using Avalonia;
using Avalonia.Controls;
using OfficeIMO.IWork;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed partial class StudioIWorkConversionTests {
    [Theory]
    [InlineData("simple.pages", "pages-pdf", 1, "hello pages", 840)]
    [InlineData("simple.pages", "pages-pdf", 1, "hello pages", 1280)]
    [InlineData("numbers-unpadded-scalar.numbers", "numbers-pdf", 1, "42", 840)]
    [InlineData("numbers-unpadded-scalar.numbers", "numbers-pdf", 1, "42", 1280)]
    [InlineData("complete.key", "keynote-pdf", 2, "hello keynote", 840)]
    [InlineData("complete.key", "keynote-pdf", 2, "hello keynote", 1280)]
    public async Task Explicit_PDF_choice_runs_complete_iwork_conversion_and_shows_shared_evidence(
        string fixture, string route, int pages, string text, int width) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = width > 1000 ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, fixture);
            if (fixture == "complete.key") {
                IWorkKeynoteDocument keynote = IWorkKeynoteDocument.Create();
                keynote.AddSlide().AddText("hello keynote", 40, 40, 600, 180, fontSizePoints: 24);
                keynote.AddSlide().AddText("Second slide", 40, 40, 600, 180, fontSizePoints: 24);
                keynote.Save(source);
            } else {
                File.Copy(Fixture(fixture), source);
            }
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var queue = model.ConversionWorkbench;
            queue.SelectedRoute = queue.Routes.Single(choice => choice.Route.Id == route);
            Assert.True(queue.AddDroppedPaths([source]));
            ConversionJobViewModel job = Assert.Single(queue.Jobs);
            Assert.Equal(route, job.Route.Route.Id);
            Assert.False(job.SupportsIWorkOptions); // The PDF route enforces complete editable content.
            var view = new ConversionWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = width > 1000 ? 800 : 600, Content = view };
            try {
                window.Show(); window.UpdateLayout();
                await queue.RunQueueCommand.ExecuteAsync(null);
                Assert.Equal(ConversionJobState.Completed, job.State);
                PdfReadDocument pdf = PdfReadDocument.Open(File.ReadAllBytes(job.OutputPath!));
                Assert.Equal(pages, pdf.Pages.Count);
                Assert.Contains(text, pdf.ExtractText());
                Assert.Equal("iwork-editable,pdf", job.ConversionEvidence!.Facts["conversionStages"]);
                Assert.Equal("False", job.ConversionEvidence.Facts["partialEditableReconstruction"]);
                Assert.Equal(64, job.SourceFingerprint.Length);
                Assert.Contains(job.Diagnostics, diagnostic => diagnostic.Code == "OutputReopened");
                await queue.PreviewOutputCommand.ExecuteAsync(null);
                Assert.Equal(pages, queue.OutputPreviewPages.Count);
                view.FindControl<StackPanel>("ConversionEvidencePanel")!.BringIntoView(); window.UpdateLayout();
                Assert.True(view.FindControl<StackPanel>("ConversionEvidencePanel")!.IsEffectivelyVisible);
                Capture(window, route + "-" + width);
                return true;
            } finally { window.Close(); }
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData("simple.numbers", "numbers-pdf", "IWORK_NUMBERS_CELL_PADDING_OMITTED")]
    [InlineData("simple.key", "keynote-pdf", "speaker-notes-omitted")]
    public async Task Complete_PDF_choice_rejects_native_content_omissions(string fixture, string route, string code) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, fixture);
            File.Copy(Fixture(fixture), source);
            byte[] original = File.ReadAllBytes(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var queue = model.ConversionWorkbench;
            queue.SelectedRoute = queue.Routes.Single(choice => choice.Route.Id == route);
            Assert.True(queue.AddDroppedPaths([source]));
            ConversionJobViewModel job = Assert.Single(queue.Jobs);
            await queue.RunQueueCommand.ExecuteAsync(null);
            Assert.Equal(ConversionJobState.Failed, job.State);
            Assert.Null(job.OutputPath);
            Assert.Equal(original, File.ReadAllBytes(source));
            if (route == "numbers-pdf") {
                // Rejection during editable projection occurs before a conversion result exists.
                Assert.Contains("padding", job.Summary!, StringComparison.OrdinalIgnoreCase);
            } else {
                Assert.Contains(job.ConversionEvidence!.FidelityDiagnostics, diagnostic =>
                    diagnostic.Code == code && diagnostic.LossKind == OfficeConversionLossKind.Omission);
            }
            Assert.False(job.SupportsIWorkOptions);
            return true;
        }, CancellationToken.None);
    }
}
