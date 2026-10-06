using System;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsTiffTwelveBitTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void IndependentTwelveBitTiffsSurvivePackageAndExport(XpsFormat format) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Tiff12");
        string[] files = Directory.GetFiles(corpus, "*.tif");
        Assert.NotEmpty(files);
        foreach (string file in files) {
            byte[] bytes = File.ReadAllBytes(file);
            Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out var expected));
            var document = XpsDocument.Create(format);
            string source = document.AddResource("Images/source.tif", bytes, "image/tiff");
            document.AddPage(105, 57).AddImage(source, 0, 0, 105, 57);
            var loaded = XpsDocument.Load(document.Save());
            var page = loaded.Pages[0];
            var svg = page.ToSvg();
            Assert.Empty(svg.Diagnostics);
            var direct = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var drawing));
            var svgImage = OfficeDrawingRasterRenderer.Render(drawing!, background: OfficeColor.White);
            var pdfPage = Assert.Single(PdfReadDocument.Open(loaded.ToPdf()).Pages);
            var pdfImage = OfficeDrawingRasterRenderer.Render(pdfPage.ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                var pixel = expected!.GetPixel(x, y);
                int Composite(int value) => (value * pixel.A + 255 * (255 - pixel.A) + 127) / 255;
                foreach (var image in new[] { direct, svgImage, pdfImage }) {
                    var actual = image.GetPixel(x * 3 + 1, y * 3 + 1);
                    Assert.InRange(Math.Abs(actual.R - Composite(pixel.R)), 0, 2);
                    Assert.InRange(Math.Abs(actual.G - Composite(pixel.G)), 0, 2);
                    Assert.InRange(Math.Abs(actual.B - Composite(pixel.B)), 0, 2);
                    Assert.Equal(255, actual.A);
                }
            }
        }
    }
}
