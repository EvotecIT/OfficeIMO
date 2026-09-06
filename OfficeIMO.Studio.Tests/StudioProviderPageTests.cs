using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioProviderPageTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 820, true)]
    public async Task PreviewAndPageExportAcceptProviderDocuments(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            var provider = new TestStorageFile("content://documents/opaque-page-source", CreatePdf(), "Selected print and page export document.pdf");
            provider.BeforeRead = () => { if (provider.Reads > 1) throw new IOException("The preview read its source twice."); };
            string location = await services.Storage.RegisterAsync(provider.Item, default);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(location), services: services);
            var preview = model.OutputWorkbench.PrintPreview;
            var window = new Window { Width = width, Height = height, Content = new OutputIntakeWorkbenchView { DataContext = model } };
            try {
                window.Show();
                await preview.ChooseInputCommand.ExecuteAsync(null);
                Assert.Equal(provider.Name, preview.InputName);
                preview.SelectedPagesPerSheet = preview.PagesPerSheetChoices.Single(choice => choice.Value == 2);
                await preview.BuildPreviewCommand.ExecuteAsync(null);
                Assert.True(preview.HasPreview, preview.Status);
                Assert.Equal(2, Assert.Single(preview.Sheets).Placements.Count);
                Assert.Equal(1, provider.Reads);
                Assert.Equal(1, provider.ClosedReads);
                Capture(window, $"provider-print-{width}-{(dark ? "dark" : "light")}.png");
                provider.BeforeRead = null;
                provider.DenyRead = true;
                await preview.BuildPreviewCommand.ExecuteAsync(null);
                Assert.False(preview.HasPreview);
                Assert.Contains("permission", preview.Status, StringComparison.OrdinalIgnoreCase);
                provider.DenyRead = false;

                model.OutputWorkbench.ShowPageExportCommand.Execute(null);
                var export = model.OutputWorkbench.PageExport;
                await export.ChooseInputCommand.ExecuteAsync(null);
                Assert.Equal(provider.Name, export.InputName);
                Assert.False(export.ExportCommand.CanExecute(null));
                Assert.Empty(export.OutputDirectory);
                export.OutputDirectory = Path.Combine(services.Paths.Root, "exported-pages");
                export.MaximumDimension = 128;
                await export.ExportCommand.ExecuteAsync(null);
                Assert.True(export.HasOutput, export.Summary);
                string[] images = Directory.GetFiles(export.PublishedDirectory!);
                Assert.Equal(2, images.Length);
                Assert.All(images, path => Assert.True(OfficeImageReader.TryValidateContent(File.ReadAllBytes(path), path, default, out _)));
                Assert.Equal(0, provider.Writes);
                Assert.Equal(provider.Reads - 1, provider.ClosedReads); // One denied request never returned a stream.
                Capture(window, $"provider-pages-{width}-{(dark ? "dark" : "light")}.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static byte[] CreatePdf() => PdfDocument.Create(compose => {
        for (int page = 1; page <= 2; page++) {
            string text = "Provider page " + page;
            compose.Page(sheet => sheet.Size(300, 400).Content(content => content.Item(item =>
                item.Paragraph(paragraph => paragraph.Text(text)))));
        }
    }).ToBytes();

    private static void Capture(Window window, string name) {
        window.UpdateLayout();
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
