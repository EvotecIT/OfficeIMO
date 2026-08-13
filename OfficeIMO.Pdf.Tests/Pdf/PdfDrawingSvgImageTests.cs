using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingSvgImageTests {
    [Fact]
    public void DrawingSvgImageIsRasterizedIntoPdf() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10'><rect width='10' height='10' fill='red'/><rect x='10' width='10' height='10' fill='blue'/></svg>");
        var projection = new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 40D, 20D));
        OfficeDrawing drawing = new OfficeDrawing(40D, 20D)
            .AddImage(svg, "image/svg+xml", projection, "Red and blue status image");

        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Drawing(drawing)
            .ToBytes();

        Assert.Equal("%PDF", Encoding.ASCII.GetString(pdf, 0, 4));
        PdfExtractedImage extracted = Assert.Single(PdfDocument.Open(pdf).Read.Images());
        Assert.True(OfficeRasterImageDecoder.TryDecode(extracted.Bytes, out OfficeRasterImage? raster));
        Assert.NotNull(raster);
        byte[] pixels = raster!.GetPixels().ToArray();
        int left = (raster.Height / 2 * raster.Width + raster.Width / 4) * 4;
        int right = (raster.Height / 2 * raster.Width + raster.Width * 3 / 4) * 4;
        Assert.True(pixels[left] > pixels[left + 2]);
        Assert.True(pixels[right + 2] > pixels[right]);
    }
}
