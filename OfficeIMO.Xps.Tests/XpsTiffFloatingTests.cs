using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsTiffFloatingTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "Fixtures", "TiffFloating");

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void IndependentlyEncodedFloatingTiffsRetainPaintAcrossRasterSvgAndPdf(XpsFormat format) {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            byte[] expected = File.ReadAllBytes(Path.Combine(Corpus, Path.ChangeExtension(fields[0], ".rgba")));
            var document = XpsDocument.Create(format);
            string source = document.AddResource("Images/source.tif", File.ReadAllBytes(Path.Combine(Corpus, fields[0])), "image/tiff");
            document.AddPage(57, 51).AddImage(source, 0, 0, 57, 51);
            if (fields[0].Contains("photo5", StringComparison.Ordinal)) {
                var originalPage = document.Pages[0];
                Assert.Throws<NotSupportedException>(() => originalPage.ToDrawing());
                byte[] profileBytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus", "littlecms-cmyk-lut.icc"));
                Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out var profile));
                string profileUri = document.AddResource("Profiles/source.icc", profileBytes, "application/vnd.ms-color.iccprofile");
                var markup = originalPage.GetMarkup();
                markup.Descendants().Single(e => e.Name.LocalName == "ImageBrush").SetAttributeValue("ImageSource",
                    "{ColorConvertedBitmap " + source + " " + profileUri + "}");
                originalPage.ReplaceMarkup(markup);
                for (int y = 0; y < 17; y++) for (int x = 0; x < 19; x++) {
                    double alpha = ((x + y) % 5) / 4D;
                    double[] channels = Enumerable.Range(0, 4).Select(c =>
                        fields[0].Contains("a1", StringComparison.Ordinal) && alpha == 0 ? 0D : ((x * 3 + y * 5 + c * 7) % 17) / 16D).ToArray();
                    Assert.True(profile!.TryConvert(channels, OfficeIccRenderingIntent.RelativeColorimetric, out var converted));
                    int offset = (y * 19 + x) * 4;
                    expected[offset] = converted.R; expected[offset + 1] = converted.G; expected[offset + 2] = converted.B;
                }
            }
            var page = XpsDocument.Load(document.Save()).Pages[0];
            var svg = page.ToSvg();
            Assert.Empty(svg.Diagnostics);
            var direct = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var svgDrawing));
            var svgImage = OfficeDrawingRasterRenderer.Render(svgDrawing!, scale: 1, background: OfficeColor.White);
            var pdfPage = Assert.Single(PdfReadDocument.Open(document.ToPdf()).Pages);
            var pdfImage = OfficeDrawingRasterRenderer.Render(pdfPage.ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
            for (int y = 0; y < 17; y++) for (int x = 0; x < 19; x++) {
                int offset = (y * 19 + x) * 4;
                foreach (var (image, lane) in new[] { (direct, "direct"), (svgImage, "svg"), (pdfImage, "pdf") }) {
                    var actual = image.GetPixel(x * 3 + 1, y * 3 + 1);
                    int Composite(int channel) => (expected[offset + channel] * expected[offset + 3] +
                        255 * (255 - expected[offset + 3]) + 127) / 255;
                    Assert.True(Math.Abs(actual.R - Composite(0)) <= 3, $"{fields[0]} {lane} {x},{y}: R {actual.R}, expected {Composite(0)}; dimensions {image.Width}x{image.Height}");
                    Assert.InRange(Math.Abs(actual.G - Composite(1)), 0, 3);
                    Assert.InRange(Math.Abs(actual.B - Composite(2)), 0, 3);
                    Assert.Equal(255, actual.A);
                }
            }
        }
    }
}
