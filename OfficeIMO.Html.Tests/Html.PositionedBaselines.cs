using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPositionedBaselineTests {
    [Fact]
    public void SavedPdfAndSvgUseTheSameAuthoredTextBaselines() {
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, Margins = HtmlRenderMargins.All(0), DefaultFontFamily = "Pinned" };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf")));
        var document = HtmlConversionDocument.Parse("<p style='margin:16px;font:20px/32px Pinned'>A<sub>2</sub>B<sup>3</sup>C</p>");
        string svg = System.Text.Encoding.UTF8.GetString(document.ExportImages(OfficeImageExportFormat.Svg, options)[0].Bytes);
        XElement[] nodes = XDocument.Parse(svg).Descendants().Where(node => node.Name.LocalName == "text").ToArray();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options)));
        var page = pdf.GetPage(1);
        foreach (XElement node in nodes) {
            var letter = Assert.Single(page.Letters, letter => letter.Value == node.Value);
            double expected = double.Parse(node.Attribute("y")!.Value, CultureInfo.InvariantCulture);
            double actual = ((double)page.Height - letter.StartBaseLine.Y) / 0.75D;
            Assert.InRange(actual - expected, -0.02D, 0.02D);
        }
    }

    [Fact]
    public void TopLevelInlineContentDoesNotReportAChangeOfPageGeometry() {
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, PageSize = new OfficePageSize(5, 4), Margins = HtmlRenderMargins.All(0) };
        var result = HtmlConversionDocument.Parse("<span>First inline line</span><br><svg width='100' height='30'><text x='0' y='20'>SVG label</text></svg>")
            .ExportImages(OfficeImageExportFormat.Png, options);
        Assert.Single(result);
        Assert.DoesNotContain(result[0].Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
    }
}
