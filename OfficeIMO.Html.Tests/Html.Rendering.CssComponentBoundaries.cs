using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlCssComponentBoundaryTests {
    [Fact]
    public void ComponentBoundaryFixturePreservesTextAndAnExplicitPageBreak() {
        string html = System.IO.File.ReadAllText(System.IO.Path.Combine(AppContext.BaseDirectory, "Documents", "Html", "Css", "component-boundaries.html"));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions()));
        Assert.Equal(2, pdf.NumberOfPages);
        string first = pdf.GetPage(1).Text;
        Assert.Contains("Curly block", first);
        Assert.Contains("Square block", first);
        Assert.Contains("Nested blocks", first);
        Assert.Contains("Escaped delimiter", first);
        Assert.Contains("Page break control", pdf.GetPage(2).Text);
    }

    [Theory]
    [InlineData("{ignored;display:none;}")]
    [InlineData("[ignored;display:none;]")]
    [InlineData("{[ignored;display:none;]}")]
    [InlineData("[ignored\\];display:none;]")]
    public void NestedCustomPropertyComponentsDoNotBecomeRenderedDeclarations(string payload) {
        string html = "<div id='kept' style='--payload:" + payload + ";width:120px;height:40px;background:#006699;color:white'>VISIBLE</div>";
        var document = HtmlConversionDocument.Parse(html);
        var options = new HtmlRenderOptions { ViewportWidth = 200, ViewportHeight = 90, Margins = HtmlRenderMargins.All(0) };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        HtmlRenderShape shape = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), item => item.Source == "div#kept" && item.Shape.FillColor.HasValue);
        Assert.Equal(OfficeColor.FromRgb(0, 102, 153), shape.Shape.FillColor);
        Assert.Contains("#006699", document.ToSvg(options), StringComparison.OrdinalIgnoreCase);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions()));
        Assert.Contains("VISIBLE", string.Concat(pdf.GetPages().Select(page => page.Text)));
    }
}
