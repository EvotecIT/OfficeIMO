using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlContainingHeightTests {
    [Theory]
    [InlineData("inline", "height:240px;max-height:200%", "", 240D, 240D)]
    [InlineData("contents", "height:240px;max-height:200%", "", 240D, 240D)]
    [InlineData("inline", "height:200%;line-height:30px", "X", 30D, 30D)]
    [InlineData("contents", "height:200%;line-height:30px", "X", 30D, 30D)]
    [InlineData("inline", "height:1px;min-height:200%", "", 1D, 20D)]
    [InlineData("contents", "height:1px;min-height:200%", "", 1D, 20D)]
    public void OrdinaryWrappersKeepAnIndefiniteAncestorIndefinite(string display, string constraints,
        string text, double expectedHeight, double expectedFollowing) {
        HtmlRenderDocument rendered = Render("<div><span style='display:" + display + ";height:20px'>"
            + Fill(constraints, text) + "</span></div>" + Following);

        Assert.Equal(expectedHeight, Shape(rendered, "div#fill").Height, 6);
        Assert.Equal(expectedFollowing, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("block", "inline")]
    [InlineData("block", "contents")]
    [InlineData("inline-block", "inline")]
    [InlineData("inline-block", "contents")]
    public void DefiniteBlockContentHeightPassesThroughOrdinaryWrappers(string parentDisplay, string wrapperDisplay) {
        HtmlRenderDocument rendered = Render("<div style='display:" + parentDisplay + ";height:120px;width:200px;vertical-align:top'>"
            + "<span style='display:" + wrapperDisplay + ";height:20px'>" + Fill("height:100%") + "</span></div>" + Following);

        Assert.Equal(120D, Shape(rendered, "div#fill").Height, 6);
        Assert.Equal(120D, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("block", "20px", "height:100%", "", 20D)]
    [InlineData("inline-block", "20px", "height:100%", "", 20D)]
    [InlineData("block", "auto", "height:200%;line-height:30px", "X", 30D)]
    [InlineData("inline-block", "auto", "height:200%;line-height:30px", "X", 30D)]
    public void ActualFormattingBoxesStartTheirOwnHeightBasis(string display, string height,
        string constraints, string text, double expectedHeight) {
        HtmlRenderDocument rendered = Render("<div style='height:120px'><div style='display:" + display
            + ";height:" + height + ";width:200px;vertical-align:top'><span style='height:70px'>"
            + Fill(constraints, text) + "</span></div></div>" + Following);

        Assert.Equal(expectedHeight, Shape(rendered, "div#fill").Height, 6);
        Assert.Equal(120D, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("content-box", "inline", 120D, 144D)]
    [InlineData("content-box", "contents", 120D, 144D)]
    [InlineData("border-box", "inline", 96D, 120D)]
    [InlineData("border-box", "contents", 96D, 120D)]
    public void ForwardedContentHeightDoesNotDeductWrapperDecoration(string sizing, string display,
        double expectedHeight, double expectedFollowing) {
        HtmlRenderDocument rendered = Render("<div style='height:120px;padding:10px;border:2px solid red;box-sizing:" + sizing + "'>"
            + "<span style='display:" + display + ";height:20px;padding:7px;border:1px solid black;box-sizing:border-box'>"
            + Fill("height:100%") + "</span></div>" + Following);

        HtmlRenderShape fill = Shape(rendered, "div#fill");
        Assert.Equal(expectedHeight, fill.Height, 6);
        Assert.Equal(12D, fill.Y, 6);
        Assert.Equal(expectedFollowing, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void ReplacedInlineKeepsItsOwnAuthoredHeight() {
        const string pixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg==";
        HtmlRenderDocument rendered = Render("<div style='height:120px'><span style='height:70px'>"
            + "<img id='fill' style='display:inline;height:20px;width:20px;vertical-align:top' src='data:image/png;base64,"
            + pixel + "'></span></div>" + Following);

        Assert.Equal(20D, Assert.Single(Visuals(rendered).OfType<HtmlRenderImage>()).Height, 6);
        Assert.Equal(120D, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    private const string Following = "<div id='following' style='height:10px;background:blue'></div>";
    private static string Fill(string constraints, string text = "") =>
        "<div id='fill' style='display:inline-block;vertical-align:top;width:40px;background:lime;" + constraints + "'>" + text + "</div>";

    private static HtmlRenderDocument Render(string content) {
        var options = new HtmlRenderOptions { ViewportWidth = 600D, ViewportHeight = 900D,
            PageSize = new OfficePageSize(600D / 96D, 900D / 96D), Margins = HtmlRenderMargins.All(0D),
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser, HonorCssPageRules = false };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
        string html = "<!doctype html><style>html,body{margin:0;padding:0;font:16px/20px Pinned}</style>" + content;
        return HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: options)).Document;
    }

    private static HtmlRenderShape Shape(HtmlRenderDocument rendered, string source) => Assert.Single(
        Visuals(rendered).OfType<HtmlRenderShape>(), shape => shape.Source == source && shape.Shape.FillColor.HasValue);
    private static IEnumerable<HtmlRenderVisual> Visuals(HtmlRenderDocument rendered) => rendered.Pages.SelectMany(page => Enumerate(page.Scene));
    private static IEnumerable<HtmlRenderVisual> Enumerate(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderSemanticGroup group => group.Visuals,
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                _ => null
            };
            if (children != null) foreach (HtmlRenderVisual child in Enumerate(children)) yield return child;
        }
    }
}
