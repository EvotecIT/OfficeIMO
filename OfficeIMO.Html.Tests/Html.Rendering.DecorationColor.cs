using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRendering_TextDecorationColorSurvivesSceneSvgAndPdfBackends() {
        const string html = "<p><span style='color:#0000ff;text-decoration-line:underline line-through;"
            + "text-decoration-style:dashed;text-decoration-color:#ff0000'>DecoratedPdf</span></p>";
        var renderOptions = new HtmlRenderOptions {
            ViewportWidth = 220D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(10D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, renderOptions);
        HtmlRenderText text = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            visual => visual.Text == "DecoratedPdf");
        string svg = Encoding.UTF8.GetString(
            HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, renderOptions).Bytes);
        var pdfOptions = new HtmlToPdfOptions(renderOptions) {
            PdfOptions = new OfficeIMO.Pdf.PdfOptions { CompressContentStreams = false }
        };
        string pdf = Encoding.ASCII.GetString(HtmlConversionDocument.Parse(html).ToPdfBytes(pdfOptions));

        Assert.Equal(OfficeColor.Blue, text.Color);
        Assert.Equal(OfficeColor.Red, text.DecorationColor);
        Assert.Equal(OfficeTextDecorationStyle.Dashed, text.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Dashed, text.StrikethroughStyle);
        Assert.Contains("text-decoration-color=\"#FF0000\"", svg, StringComparison.Ordinal);
        Assert.Contains("1 0 0 RG", pdf, StringComparison.Ordinal);
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(text-decoration-color:#ff0000)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(text-decoration-color:paint(accent))"));
    }
}
