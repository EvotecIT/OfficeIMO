using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsImageDefaultsTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void NativeIntegerImagesRetainSampleColorsAlphaAndTiffSampleOrder(XpsFormat format) {
        string fixtures = Path.Combine(AppContext.BaseDirectory, "Fixtures", "ImageDefaults");
        foreach (var group in File.ReadAllLines(Path.Combine(fixtures, "expected.csv")).Skip(1).GroupBy(row => row.Split(',')[0])) {
            var document = XpsDocument.Create(format);
            string source = document.AddResource("Images/source" + Path.GetExtension(group.Key), File.ReadAllBytes(Path.Combine(fixtures, group.Key)),
                group.Key.EndsWith(".tif", StringComparison.Ordinal) ? "image/tiff" : "image/jpeg");
            document.AddPage(16, 8).AddImage(source, 0, 0, 16, 8);
            var page = XpsDocument.Load(document.Save()).Pages[0];
            if (group.Key.Contains("oriented"))
                Assert.Equal("0,0,16,8", (string?)page.GetMarkup().Descendants().Single(e => e.Name.LocalName == "ImageBrush").Attribute("Viewbox"));
            Assert.Empty(page.ToSvg().Diagnostics);
            Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png,
                new OfficeImageExportOptions { BackgroundColor = OfficeColor.Transparent }).Bytes, out var image));
            foreach (string row in group) {
                string[] values = row.Split(',');
                var pixel = image!.GetPixel(int.Parse(values[1]), int.Parse(values[2]));
                Assert.InRange(Math.Abs(pixel.R - int.Parse(values[3])), 0, 3);
                Assert.InRange(Math.Abs(pixel.G - int.Parse(values[4])), 0, 3);
                Assert.InRange(Math.Abs(pixel.B - int.Parse(values[5])), 0, 3);
                Assert.Equal(int.Parse(values[6]), (int)pixel.A);
            }
            Assert.Single(PdfReadDocument.Open(document.ToPdf()).Pages);
        }
    }
}
