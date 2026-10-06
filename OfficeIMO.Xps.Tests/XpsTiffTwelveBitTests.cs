using System;
using System.IO;
using System.Linq;
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
            byte[]? nativeIcc = null;
            if (Path.GetFileName(file).StartsWith("k6-", StringComparison.Ordinal)) {
                Assert.Throws<NotSupportedException>(() => document.Pages[0].ToDrawing());
                byte[] profileBytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus", "littlecms-cmyk-lut.icc"));
                nativeIcc = File.ReadAllBytes(file + ".srgb");
                string profileUri = document.AddResource("Profiles/source.icc", profileBytes, "application/vnd.ms-color.iccprofile");
                var markup = document.Pages[0].GetMarkup();
                markup.Descendants().Single(e => e.Name.LocalName == "ImageBrush").SetAttributeValue("ImageSource",
                    "{ColorConvertedBitmap " + source + " " + profileUri + "}");
                document.Pages[0].ReplaceMarkup(markup);
            }
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
                int at = (y * 35 + x) * 3;
                var pixel = nativeIcc == null ? expected!.GetPixel(x, y) : OfficeColor.FromRgb(nativeIcc[at], nativeIcc[at + 1], nativeIcc[at + 2]);
                int tolerance = nativeIcc == null ? 2 : 5; // Native ICC rounding plus export quantization.
                int Composite(int value) => (value * pixel.A + 255 * (255 - pixel.A) + 127) / 255;
                foreach (var image in new[] { direct, svgImage, pdfImage }) {
                    var actual = image.GetPixel(x * 3 + 1, y * 3 + 1);
                    Assert.InRange(Math.Abs(actual.R - Composite(pixel.R)), 0, tolerance);
                    Assert.InRange(Math.Abs(actual.G - Composite(pixel.G)), 0, tolerance);
                    Assert.InRange(Math.Abs(actual.B - Composite(pixel.B)), 0, tolerance);
                    Assert.Equal(255, actual.A);
                }
            }
        }
    }
}
