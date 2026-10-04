using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Assistant;

namespace OfficeIMO.Studio.Tests;

public sealed partial class DocumentAssistantTests {
    [Theory]
    [InlineData(440, 720, false)]
    [InlineData(960, 620, true)]
    public async Task MixedPageOffersOcrAndDisclosesIncompleteEvidence(int width, int height, bool light) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            services.AiConnections.ProviderIndex = 3; services.AiConnections.Model = "fixture"; services.AiConnections.IsConnected = true;
            using var raster = new RenderTargetBitmap(new PixelSize(20, 20));
            using (var drawing = raster.CreateDrawingContext()) drawing.DrawRectangle(Brushes.Gray, null, new Rect(0, 0, 20, 20));
            using var png = new MemoryStream(); raster.Save(png, PngBitmapEncoderOptions.Default);
            var pdf = PdfDocument.Create(new PdfOptions { PageWidth = 300, PageHeight = 220, MarginLeft = 24, MarginRight = 24, MarginTop = 24, MarginBottom = 24 });
            pdf.Content.Image(png.ToArray(), 180, 120);
            pdf.Content.Paragraph(paragraph => paragraph.Text("Footer 1"));
            bool current = true, opened = false;
            using var model = new DocumentAssistantViewModel(services.AiConnections, _ => new(pdf.ToBytes(), "mixed.pdf", null, () => current),
                _ => { }, services.Localizer, (_, _, _) => throw new InvalidOperationException("Readiness must not connect")) {
                OpenOcr = _ => { opened = true; return Task.CompletedTask; }
            };
            await model.PrepareEvidenceCommand.ExecuteAsync(null);
            Assert.True(model.PreparedDocument!.HasSourceDiagnostics);
            Assert.True(model.CanOpenOcr);
            var view = new AssistantView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view, RequestedThemeVariant = light ? ThemeVariant.Light : ThemeVariant.Dark };
            try {
                window.Show(); window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                var button = Assert.Single(view.GetVisualDescendants().OfType<Button>(), item => ReferenceEquals(item.Command, model.OpenOcrCommand));
                Assert.True(button.IsVisible);
                var position = button.TranslatePoint(default, window)!.Value;
                Assert.InRange(position.Y, 0, height - button.Bounds.Height + 1);
                using var frame = window.CaptureRenderedFrame();
                Assert.NotNull(frame);
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (output is not null) {
                    Directory.CreateDirectory(output);
                    frame.Save(Path.Combine(output, $"assistant-mixed-ocr-{width}x{height}-{(light ? "light" : "dark")}.png"), PngBitmapEncoderOptions.Default);
                }
                await model.OpenOcrCommand.ExecuteAsync(null);
                Assert.True(opened);
                opened = false; current = false;
                Assert.False(model.CanOpenOcr);
                await model.OpenOcrCommand.ExecuteAsync(null);
                Assert.False(opened);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
