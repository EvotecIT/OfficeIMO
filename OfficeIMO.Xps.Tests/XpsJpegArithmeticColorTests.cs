using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsJpegArithmeticColorTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void ProfiledCmykAndYcckPreserveNativeColorsThroughExport(XpsFormat format) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Fixtures", "JpegArithmeticColor");
        byte[] profile = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus", "littlecms-cmyk-lut.icc"));
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string name = row.Split(',')[0];
            byte[] native = File.ReadAllBytes(Path.Combine(corpus, name + ".srgb"));
            var document = XpsDocument.Create(format);
            string source = document.AddResource("Images/source.jpg", File.ReadAllBytes(Path.Combine(corpus, name)), "image/jpeg");
            string profileUri = document.AddResource("Profiles/source.icc", profile, "application/vnd.ms-color.iccprofile");
            var page = document.AddPage(105, 57).AddImage(source, 0, 0, 105, 57);
            var markup = page.GetMarkup();
            markup.Descendants().Single(e => e.Name.LocalName == "ImageBrush").SetAttributeValue("ImageSource",
                "{ColorConvertedBitmap " + source + " " + profileUri + "}");
            page.ReplaceMarkup(markup);
            var loaded = XpsDocument.Load(document.Save()); page = loaded.Pages[0];
            var svg = page.ToSvg(); Assert.Empty(svg.Diagnostics);
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), out var drawing));
            var images = new[] {
                OfficeDrawingRasterRenderer.Render(page.ToDrawing()),
                OfficeDrawingRasterRenderer.Render(drawing!),
                OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(loaded.ToPdf()).Pages[0].ToDrawing(), scale: 4D / 3D)
            };
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                int at = (y * 35 + x) * 3;
                foreach (var image in images) {
                    var pixel = image.GetPixel(x * 3 + 1, y * 3 + 1);
                    Assert.True(Math.Abs(pixel.R - native[at]) <= 5 && Math.Abs(pixel.G - native[at + 1]) <= 5 &&
                        Math.Abs(pixel.B - native[at + 2]) <= 5 && pixel.A == 255, $"{name} at {x},{y}");
                }
            }
        }
    }
}
