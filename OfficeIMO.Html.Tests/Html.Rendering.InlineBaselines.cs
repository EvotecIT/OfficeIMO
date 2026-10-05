using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("display:block")]
    [InlineData("display:flex;flex-direction:row;align-items:flex-start")]
    [InlineData("display:flex;flex-direction:column;align-items:flex-start")]
    public void HtmlInlineText_MixedSizesShareTheSavedPdfAndSvgBaseline(string containerStyle) {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            Margins = HtmlRenderMargins.All(0D),
            DefaultFontFamily = "Pinned"
        };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf")));
        var document = HtmlConversionDocument.Parse("<div style='" + containerStyle + ";width:600px;font:40px Pinned'>"
            + "<span>AAAA <span style='font-size:10px'>BBBB</span></span></div>");
        string svg = document.ToSvg(options);
        XElement[] text = XDocument.Parse(svg).Descendants().Where(node => node.Name.LocalName == "text").ToArray();
        XElement large = Assert.Single(text, node => node.Value.Trim() == "AAAA");
        XElement small = Assert.Single(text, node => node.Value.Trim() == "BBBB");
        Assert.Equal(ParseY(large), ParseY(small), 3);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options)));
        var letters = pdf.GetPage(1).Letters;
        double largeBaseline = letters.First(letter => letter.Value == "A").StartBaseLine.Y;
        double smallBaseline = letters.First(letter => letter.Value == "B").StartBaseLine.Y;
        Assert.Equal(largeBaseline, smallBaseline, 3);

        static double ParseY(XElement node) => double.Parse(node.Attribute("y")!.Value, CultureInfo.InvariantCulture);
    }
}
