using Avalonia;
using Avalonia.Controls;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed partial class StudioIWorkConversionTests {
    [Theory]
    [InlineData("simple.pages", "pages-pdf", 1, "hello pages", 840)]
    [InlineData("simple.pages", "pages-pdf", 1, "hello pages", 1280)]
    [InlineData("simple.numbers", "numbers-pdf", 1, "Z", 840)]
    [InlineData("simple.numbers", "numbers-pdf", 1, "Z", 1280)]
    [InlineData("simple.key", "keynote-pdf", 2, "hello keynote", 840)]
    [InlineData("simple.key", "keynote-pdf", 2, "hello keynote", 1280)]
    public async Task Explicit_PDF_choice_runs_complete_iwork_conversion_and_shows_shared_evidence(
        string fixture, string route, int pages, string text, int width) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = width > 1000 ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, fixture);
            File.Copy(Fixture(fixture), source);
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
}
