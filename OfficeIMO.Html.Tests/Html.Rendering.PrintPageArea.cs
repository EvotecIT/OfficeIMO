using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPdf_FixedPercentageHeaderUsesPageAreaWithoutExpandingPrintLayout() {
        const string html = "<style>@page{size:400px 300px;margin:40px}html,body{margin:0}</style>"
            + "<div id='fixed' style='position:fixed;left:0;top:0;width:100%;height:25px;background:blue'></div>"
            + "<div style='height:250px'>Flow</div>";
        var result = HtmlConversionDocument.Parse(html).RenderToPdfResult(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions()));

        Assert.Equal(2, result.RenderResult.Document.Pages.Count);
        foreach (var page in result.RenderResult.Document.Pages) {
            Assert.Equal(400D, page.Width, 3);
            var shape = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(), v => v.Source == "div#fixed");
            Assert.Equal(40D, shape.X, 3);
            Assert.Equal(40D, shape.Y, 3);
            Assert.Equal(320D, shape.Width, 3);
        }
        Assert.DoesNotContain(result.RenderResult.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.PrintFitOverflowUnresolved);
    }

    [Fact]
    public void HtmlFixedPosition_OpposingInsetsUseEachPagesAreaAndKeepStaticAnchors() {
        const string html = "<style>@page{size:400px 300px;margin:40px}@page:first{size:300px 250px;margin:20px}html,body{margin:0}</style>"
            + "<div id='edges' style='position:fixed;left:5px;right:7px;top:9px;bottom:11px;background:blue'></div>"
            + "<div style='height:30px'></div><div id='auto' style='position:fixed;width:10px;height:10px;background:red'></div>"
            + "<div style='height:150px;break-after:page'></div><div>Next</div>";
        var render = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, AutoFitWidePrintRoot = false
        });
        Assert.True(render.Pages.Count >= 2);
        foreach (var page in render.Pages) {
            var edge = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(), v => v.Source == "div#edges");
            Assert.Equal(page.Margins.Left + 5D, edge.X, 3);
            Assert.Equal(page.Margins.Top + 9D, edge.Y, 3);
            Assert.Equal(page.Width - page.Margins.Left - page.Margins.Right - 12D, edge.Width, 3);
            Assert.Equal(page.Height - page.Margins.Top - page.Margins.Bottom - 20D, edge.Height, 3);
            var automatic = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(), v => v.Source == "div#auto");
            Assert.Equal(page.Margins.Left, automatic.X, 3);
            Assert.Equal(page.Margins.Top + 30D, automatic.Y, 3);
        }
    }

    [Theory]
    [InlineData("html", "hidden")]
    [InlineData("body", "hidden")]
    [InlineData("html", "clip")]
    [InlineData("body", "clip")]
    public void HtmlPdf_AutomaticPrintFitIncludesRootScrollOverflow(string root, string overflow) {
        string html = "<style>@page{size:400px 300px;margin:0}html,body{margin:0}" + root + "{overflow-x:" + overflow + "}</style>"
            + "<div style='width:600px;height:100px;background:blue'></div>";
        var result = HtmlConversionDocument.Parse(html).RenderToPdfResult(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions()));
        Assert.Equal(600D, result.RenderResult.Document.Pages[0].Width, 3);
        Assert.DoesNotContain(result.RenderResult.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.PrintFitOverflowUnresolved);
    }

    [Theory]
    [InlineData("flex;flex-direction:column")]
    [InlineData("grid")]
    public void HtmlPdf_AutomaticPrintFitIncludesSpecializedRootScrollOverflow(string display) {
        string html = "<style>@page{size:400px 300px;margin:0}html,body{margin:0}body{display:" + display + ";max-width:400px;overflow-x:hidden}</style>"
            + "<div style='width:600px;height:100px;background:blue'></div>";
        var result = HtmlConversionDocument.Parse(html).RenderToPdfResult(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions()));
        Assert.Equal(600D, result.RenderResult.Document.Pages[0].Width, 3);
    }

    [Fact]
    public void HtmlFixedPosition_RecomputesViewportGeometryWhenDifferentSheetsHaveEqualPageAreas() {
        const string html = "<style>@page{size:400px 350px;margin:70px}@page:first{size:300px 250px;margin:20px}html,body{margin:0}</style>"
            + "<div id='fixed' style='position:fixed;left:10vw;top:10vh;width:50vw;height:20vh;background:blue'></div>"
            + "<div style='height:50px;break-after:page'>First</div><div>Second</div>";
        var render = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, AutoFitWidePrintRoot = false
        });
        var first = Assert.Single(render.Pages[0].Visuals.OfType<HtmlRenderShape>(), v => v.Source == "div#fixed");
        var later = Assert.Single(render.Pages[^1].Visuals.OfType<HtmlRenderShape>(), v => v.Source == "div#fixed");
        Assert.Equal(50D, first.X, 3);
        Assert.Equal(45D, first.Y, 3);
        Assert.Equal(150D, first.Width, 3);
        Assert.Equal(50D, first.Height, 3);
        Assert.Equal(110D, later.X, 3);
        Assert.Equal(105D, later.Y, 3);
        Assert.Equal(200D, later.Width, 3);
        Assert.Equal(70D, later.Height, 3);
    }
}
