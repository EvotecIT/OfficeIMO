using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSvgSurfaceRoundingTests {
    [Fact]
    public void FractionalDpiSvgDimensionsMatchTheRoundedEncodedSurface() {
        byte[] pdf = PdfDocument.Create(new PdfOptions { PageWidth = 612, PageHeight = 792 })
            .Paragraph(p => p.Text("SVG surface rounding"))
            .ToBytes();
        var page = PdfReadDocument.Open(pdf).Pages[0];
        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Svg,
            new PdfImageExportOptions { TargetDpi = 150 });
        var svg = XElement.Parse(Encoding.UTF8.GetString(result.Bytes));
        Assert.Equal(1275, result.Width); Assert.Equal(1650, result.Height);
        Assert.Equal("1275px", svg.Attribute("width")?.Value);
        Assert.Equal("1650px", svg.Attribute("height")?.Value);
    }
}
