using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlVerticalTableHeaderTests {
    [Theory]
    [InlineData("block", "height:100%")]
    [InlineData("inline-block", "height:100%")]
    [InlineData("block", "height:1px;min-height:100%")]
    [InlineData("block", "height:240px;max-height:100%")]
    public void HtmlRender_TableCellHeightRemainsAvailableForPercentageDescendants(string display, string sizing) {
        string html = "<table style='width:200px'><tr><td style='height:120px;padding:1px;border:0'>"
            + "<div id='fill' style='display:" + display + ";width:40px;" + sizing + ";background:red'></div>"
            + "</td></tr></table>";
        var rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html));
        var fill = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#fill" && shape.Shape.FillColor == OfficeColor.Red);

        Assert.Equal(120D, fill.Height, 3);
    }

    [Fact]
    public void HtmlRender_TableCellHeightRemainsAvailableForPercentageImageHeight() {
        string source = "data:image/svg+xml," + Uri.EscapeDataString("<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16'><rect width='16' height='16' fill='red'/></svg>");
        string html = "<table style='width:200px'><tr><td style='height:120px;padding:1px;border:0'>"
            + "<img src='" + source + "' style='height:100%;width:40px'/></td></tr></table>";
        var rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html));
        var image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());

        Assert.Equal(120D, image.Height, 3);
    }

    [Fact]
    public void HtmlRender_TableCellMinimumWidthReservesSpaceForItsColumn() {
        const string html = "<table style='width:300px'><tr><td id='wide' style='min-width:200px;background:red'>A</td><td>B</td></tr></table>";
        var rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html));
        var cell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "td#wide" && shape.Shape.FillColor == OfficeColor.Red);

        Assert.True(cell.Width >= 200D, "The minimum-width column was squeezed to " + cell.Width);
    }

    [Theory]
    [InlineData("vertical-rl")]
    [InlineData("vertical-lr")]
    [InlineData("sideways-rl")]
    [InlineData("sideways-lr")]
    public void HtmlPdf_TableCellMinimumHeightDoesNotEllipsizeVerticalHeaders(string writingMode) {
        string html = "<style>table{width:240px}th{width:34px;height:30px;padding:8px 0 7px}"
            + "th span{display:inline-block;writing-mode:" + writingMode + ";transform:rotate(180deg);"
            + "white-space:nowrap;max-height:132px;overflow:hidden;text-overflow:ellipsis;line-height:1.2}</style>"
            + "<table><thead><tr><th><span>Cairo</span></th><th><span>LosAngeles</span></th></tr></thead>"
            + "<tbody><tr><td>One</td><td>Two</td></tr></tbody></table>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Cairo", text, StringComparison.Ordinal);
        Assert.Contains("LosAngeles", text, StringComparison.Ordinal);
        Assert.Contains("One", text, StringComparison.Ordinal);
        Assert.Contains("Two", text, StringComparison.Ordinal);
    }
}
