using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsJpegXrCmykTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void CmykProfilesPreserveChannelsAndAlphaAcrossPackageSvgAndPdf(XpsFormat format) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Fixtures", "JpegXr");
        byte[] profileBytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus", "littlecms-cmyk-lut.icc"));
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out var profile));
        foreach (string row in File.ReadLines(Path.Combine(corpus, "cmyk-manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            byte[] encoded = File.ReadAllBytes(Path.Combine(corpus, fields[0]));
            int width = int.Parse(fields[1]), height = int.Parse(fields[2]);
            byte[] expected = JpegXrTestFixture.ConvertDeviceReference(
                File.ReadAllBytes(Path.Combine(corpus, Path.ChangeExtension(fields[0], ".cmyk"))), int.Parse(fields[3]), int.Parse(fields[4]), profile!);
            foreach (bool embedded in new[] { false, true }) {
                var document = XpsDocument.Create(format);
                byte[] source = embedded ? JpegXrTestFixture.WithField(encoded, 0x8773, 7, profileBytes) : encoded;
                string imageUri = document.AddResource("Images/source.jxr", source, "image/jxr");
                var page = document.AddPage(width * 3, height * 3).AddImage(imageUri, 0, 0, width * 3, height * 3);
                if (!embedded) {
                    string profileUri = document.AddResource("Profiles/source.icc", profileBytes, "application/vnd.ms-color.iccprofile");
                    var markup = page.GetMarkup();
                    markup.Descendants().Single(e => e.Name.LocalName == "ImageBrush").SetAttributeValue("ImageSource",
                        "{ColorConvertedBitmap " + imageUri + " " + profileUri + "}");
                    page.ReplaceMarkup(markup);
                }
                document = XpsDocument.Load(document.Save());
                var svg = document.Pages[0].ToSvg(); Assert.Empty(svg.Diagnostics);
                string data = (string)XDocument.Parse(svg.Svg).Descendants().Single(e => e.Name.LocalName == "image").Attribute("href")!;
                Assert.True(OfficeRasterImageDecoder.TryDecode(Convert.FromBase64String(data.Substring(data.IndexOf(',') + 1)), out var normalized));
                Assert.Equal(expected, normalized!.GetPixels());
                if (!embedded && width == 19 && fields[0].Contains("frequency", StringComparison.Ordinal)) {
                    var pdfPage = Assert.Single(PdfReadDocument.Open(document.ToPdf()).Pages);
                    var raster = OfficeDrawingRasterRenderer.Render(pdfPage.ToDrawing(), scale: 4D / 3D, background: OfficeColor.White);
                    for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                        int offset = (y * width + x) * 4;
                        int Composite(int c) => (expected[offset + c] * expected[offset + 3] + 255 * (255 - expected[offset + 3]) + 127) / 255;
                        var pixel = raster.GetPixel(x * 3 + 1, y * 3 + 1);
                        Assert.InRange(Math.Abs(pixel.R - Composite(0)), 0, 1);
                        Assert.InRange(Math.Abs(pixel.G - Composite(1)), 0, 1);
                        Assert.InRange(Math.Abs(pixel.B - Composite(2)), 0, 1);
                    }
                }
            }
        }
    }

    [Fact]
    public void UnprofiledCmykReportsItsMissingColorContract() {
        var document = XpsDocument.Create();
        byte[] encoded = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "JpegXr", "cmyk8-frequency-q32-a0.jxr"));
        string uri = document.AddResource("Images/source.jxr", encoded, "image/jxr");
        var page = document.AddPage(19, 13).AddImage(uri, 0, 0, 19, 13);
        Assert.Throws<NotSupportedException>(() => page.ToSvg());
        var svg = page.ToSvg(allowPartial: true);
        Assert.Contains(svg.Diagnostics, d => d.Contains("CMYK image requires a usable ICC profile", StringComparison.Ordinal));
    }
}
