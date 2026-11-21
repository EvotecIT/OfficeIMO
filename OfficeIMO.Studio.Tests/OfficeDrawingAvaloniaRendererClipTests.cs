using Avalonia;
using Avalonia.Headless;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

public sealed class OfficeDrawingAvaloniaRendererClipTests {
    [Fact]
    public async Task IntersectionOfTwoNonConvexPdfClipsPaintsOnlyTheTrueRegion() {
        // Two L-shapes whose intersection is the top-left and bottom-right 40 x 40 corner squares.
        const string content =
            "20 20 m 220 20 l 220 60 l 60 60 l 60 180 l 20 180 l h W n\n" +
            "220 180 m 20 180 l 20 140 l 180 140 l 180 20 l 220 20 l h W n\n" +
            "1 0 0 rg 0 0 240 200 re f";
        byte[] pdf = System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 200] /Contents 4 0 R /Resources << >> >>", "endobj",
            "4 0 obj", $"<< /Length {content.Length} >>", "stream", content, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }) + "\n");
        OfficeDrawing drawing = PdfDocument.Load(new MemoryStream(pdf)).Pages[0].ToDrawing();

        using HeadlessUnitTestSession app = TestAppBuilder.StartSession();
        OfficeRasterImage image = await app.Dispatch(() => {
            using var raster = new RenderTargetBitmap(new PixelSize(240, 200));
            using (DrawingContext context = raster.CreateDrawingContext()) {
                context.FillRectangle(Brushes.White, new Rect(0, 0, 240, 200));
                using var renderer = new OfficeDrawingAvaloniaRenderer();
                renderer.Render(context, drawing);
            }
            using var png = new MemoryStream();
            raster.Save(png);
            return OfficeRasterImageDecoder.Decode(png.ToArray());
        }, CancellationToken.None);

        Assert.True(IsRed(image.GetPixel(40, 40)));
        Assert.True(IsRed(image.GetPixel(200, 160)));
        Assert.False(IsRed(image.GetPixel(120, 100)));
        Assert.False(IsRed(image.GetPixel(40, 100)));
        int red = 0;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                if (IsRed(image.GetPixel(x, y))) red++;
            }
        }
        Assert.InRange(red, 3000, 3400);
    }

    [Fact]
    public void PushScopesUnwindsTheScopesAlreadyPushedWhenAPushFails() {
        var disposed = new List<int>();

        Assert.Throws<InvalidOperationException>(() => OfficeDrawingAvaloniaRenderer.PushScopes(3, index =>
            index < 2 ? new Scope(() => disposed.Add(index)) : throw new InvalidOperationException()));

        Assert.Equal(new[] { 1, 0 }, disposed);
    }

    private sealed class Scope(Action dispose) : IDisposable {
        public void Dispose() => dispose();
    }

    private static bool IsRed(OfficeColor color) => color.R > 200 && color.G < 80 && color.B < 80;
}
