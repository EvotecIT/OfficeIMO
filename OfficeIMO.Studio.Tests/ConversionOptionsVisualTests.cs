using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Studio.Tests;

public sealed class ConversionOptionsVisualTests {
    [Theory]
    [InlineData(840, 600)]
    [InlineData(1280, 800)]
    public async Task SelectedJobOptionsAndSavedOutputPreviewRemainAccessible(int width, int height) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = width >= 1000 ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string source = Path.Combine(services.Paths.Root, "Visual conversion source.pdf");
            PdfDocument.Create(compose => {
                compose.Page(page => page.Size(240, 320).Content(c => c.Item(i => i.Paragraph(p => p.Text("First page")))));
                compose.Page(page => page.Size(240, 320).Content(c => c.Item(i => i.Paragraph(p => p.Text("Second page for review")))));
            }).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickWorkflowFiles: _ => Task.FromResult<IReadOnlyList<string>>([source]), services: services);
            var queue = model.ConversionWorkbench;
            queue.SelectedRoute = queue.Routes.Single(route => route.Route.Id == "pdf-docx");
            await queue.AddFilesCommand.ExecuteAsync(null);
            var job = Assert.Single(queue.Jobs);
            job.PageRanges = "2";
            job.WordMode = PdfWordImportMode.VisualPages;
            job.RasterDpi = 72;
            var view = new ConversionWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show(); window.UpdateLayout();
                var details = view.FindControl<Border>("DetailsPanel")!;
                Assert.True(details.IsEffectivelyVisible);
                Assert.True(details.Bounds.Height > 100);
                var ranges = view.GetVisualDescendants().OfType<TextBox>().Single(box => box.Text == "2");
                ranges.BringIntoView(); window.UpdateLayout();
                ranges.Focus();
                Assert.True(ranges.IsEffectivelyVisible);
                Assert.True(ranges.IsFocused);
                Capture(window, "conversion-options-" + width);
                await queue.RunQueueCommand.ExecuteAsync(null);
                Assert.True(job.HasOutput, job.Summary + " " + string.Join(" ", job.Diagnostics.Select(d => d.Message)));
                Assert.False(job.CanEditOptions);
                await queue.PreviewOutputCommand.ExecuteAsync(null);
                Assert.Single(queue.OutputPreviewPages);
                Assert.Contains("saved artifact", queue.OutputPreviewStatus);
                window.UpdateLayout();
                var images = view.GetVisualDescendants().OfType<Image>().Where(image => image.Source is not null).ToArray();
                Assert.NotEmpty(images);
                images[0].BringIntoView(); window.UpdateLayout();
                Capture(window, "conversion-output-preview-" + width);
                queue.SelectedJob = null;
                Assert.Empty(queue.OutputPreviewPages);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
