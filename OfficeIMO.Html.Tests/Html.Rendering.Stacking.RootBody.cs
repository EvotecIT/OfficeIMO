using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("ALPHA ", "", "BETA\uF42B", "ALPHA BETA")]
    [InlineData("\uF42B", "", "BETA", "BETA")]
    [InlineData("ALPHA ", "left:2000px", "BETA", "BETA")]
    public void HtmlStacking_PdfLogicalOwnerAccountsForAllRenderableLayers(string first, string firstStyle, string second, string expected) {
        string html = "<p style='margin:0'><span style='position:relative;z-index:2;font-family:MissingIcon;" + firstStyle + "'>" + first
            + "</span><span style='position:relative;z-index:1;font-family:MissingIcon'>" + second + "</span></p>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { AutoFitWidePrintContent = false });
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains(expected, text);
        Assert.DoesNotContain("\uF42B", text);
        Assert.Equal(1, text.Split("BETA").Length - 1);
        if (firstStyle.Length > 0) Assert.DoesNotContain("ALPHA", text);
    }

    [Theory]
    [InlineData("p")]
    [InlineData("h1")]
    public void HtmlStacking_PromotedInlineContextsRetainPdfLogicalReadingOrder(string tag) {
        string html = "<" + tag + " style='margin:0;font-size:16px'><span style='position:relative;z-index:2'><a href='https://example.test/alpha'>ALPHA </a></span>"
            + "<span style='position:relative;z-index:1'><a href='https://example.test/beta'>BETA</a></span></" + tag + ">";
        var options = new HtmlToPdfOptions();
        options.PdfOptions.CompressContentStreams = false;
        var result = HtmlConversionDocument.Parse(html).RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));
        byte[] pdf = result.ToBytes();
        var rendered = result.RenderResult.Document;
        Assert.Contains("/ActualText", System.Text.Encoding.ASCII.GetString(pdf));
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("ALPHA BETA", text);
        Assert.Equal(1, text.Split("ALPHA").Length - 1);
        Assert.Equal(1, text.Split("BETA").Length - 1);
        Assert.Contains("ALPHA BETA", rendered.Text);
        var info = OfficeIMO.Pdf.PdfInspector.Inspect(pdf);
        Assert.Contains("https://example.test/alpha", info.LinkUris);
        Assert.Contains("https://example.test/beta", info.LinkUris);
    }

    [Fact]
    public void HtmlStacking_NestedListContextPaintsAboveFixedHeader() {
        const string html = "<style>@page{size:400px 300px;margin:0}html,body,ul,li{margin:0;padding:0}body{display:flex;flex-direction:column}"
            + "header{position:fixed;left:0;top:0;width:100%;height:60px;background:red;z-index:1}</style>"
            + "<header></header><ul style='list-style:none'><li><section style='position:relative;z-index:2;height:60px;background:blue'>Footer</section></li></ul>";
        var rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        var raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(200, 30));
    }

    [Theory]
    [InlineData("block", 1)]
    [InlineData("block", 2)]
    [InlineData("flex", 1)]
    [InlineData("flex", 2)]
    public void HtmlStacking_PagedRootFlowContextStacksAboveEarlierFixedHeader(string display, int footerZ) {
        string html = "<style>@page{size:400px 300px;margin:0}html,body{margin:0}body{display:"
            + display + ";flex-direction:column}header{position:fixed;left:0;top:0;width:100%;height:60px;background:red;z-index:1}"
            + "footer{position:relative;z-index:" + footerZ + ";height:60px;background:blue}</style><header>Header</header><footer>Footer</footer>";
        var rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        var raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(200, 30));
    }

    [Theory]
    [InlineData("", 0)]
    [InlineData("position:relative;z-index:0", 10)]
    [InlineData("opacity:.5", 10)]
    public void HtmlStacking_RootPromotionKeepsLowerAndAncestorContextsBelowFixedHeader(string ancestor, int footerZ) {
        string html = "<style>@page{size:400px 300px;margin:0}html,body{margin:0}body{display:flex;flex-direction:column}"
            + "header{position:fixed;left:0;top:0;width:100%;height:60px;background:red;z-index:1}</style>"
            + "<header></header><main style='" + ancestor + "'><footer style='position:relative;z-index:" + footerZ
            + ";height:60px;background:blue'></footer></main>";
        var rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        var raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Red, raster.GetPixel(200, 30));
    }

    [Fact]
    public void HtmlStacking_PromotedRootFlowContextRetainsAncestorClip() {
        const string html = "<style>@page{size:400px 300px;margin:0}html,body{margin:0}body{display:flex;flex-direction:column}"
            + "header{position:fixed;left:0;top:0;width:100%;height:60px;background:red;z-index:1}</style>"
            + "<header></header><main style='width:100px;overflow:hidden'><footer style='position:relative;z-index:2;width:400px;height:60px;background:blue'></footer></main>";
        var rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, AutoFitWidePrintRoot = false });
        var raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(50, 30));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(200, 30));
    }

    [Fact]
    public void HtmlStacking_PromotedFooterKeepsContextAcrossPageFragmentation() {
        const string html = "<style>@page{size:400px 300px;margin:0}html,body{margin:0}body{display:flex;flex-direction:column}"
            + "header{position:fixed;left:0;top:0;width:100%;height:60px;background:red;z-index:1}</style>"
            + "<header></header><main style='height:280px'></main><footer style='position:relative;z-index:2;height:100px;background:blue'>"
            + "<div style='height:50px'></div><a href='https://example.test/footer'>Footer link</a></footer>";
        var rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        Assert.Equal(2, rendered.Pages.Count);
        var raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[1].CreateDrawing());
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(200, 30));
        Assert.Contains("Footer link", rendered.Text);
    }
}
