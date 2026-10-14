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
        HtmlRenderDocument rendered = Render("<div style='height:120px'><span style='height:70px'>"
            + "<img id='fill' style='display:inline;height:20px;width:20px;vertical-align:top' src='data:image/png;base64,"
            + Pixel + "'></span></div>" + Following);

        Assert.Equal(20D, Assert.Single(Visuals(rendered).OfType<HtmlRenderImage>()).Height, 6);
        Assert.Equal(120D, Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("inline", "120px", 120D)]
    [InlineData("contents", "120px", 120D)]
    [InlineData("inline", "auto", 1D)]
    [InlineData("contents", "auto", 1D)]
    public void IntrinsicReplacedWidthsUseTheSameContainingHeightAsLayout(string display, string height,
        double expectedSize) {
        HtmlRenderDocument rendered = Render(ShrinkToFit("height:" + height,
            "<span style='display:" + display + ";height:20px'>" + SquareImage("100%") + "</span>"));

        AssertReplacedContribution(rendered, expectedSize, expectedSize);
    }

    [Theory]
    [InlineData("inline")]
    [InlineData("contents")]
    public void NestedAtomicReplacedContributionsRetainTheirOwnContainingHeight(string display) {
        HtmlRenderDocument rendered = Render(ShrinkToFit("height:120px",
            "<span style='display:inline-block;height:10px;vertical-align:top'><span style='display:"
            + display + ";height:20px'>" + SquareImage("100%") + "</span></span>"));

        AssertReplacedContribution(rendered, 10D, 10D);
    }

    [Theory]
    [InlineData("inline-flex", false, false)]
    [InlineData("inline-flex", true, false)]
    [InlineData("inline-grid", false, false)]
    [InlineData("inline-grid", true, false)]
    [InlineData("inline-flex", false, true)]
    [InlineData("inline-grid", false, true)]
    public void BlockifiedItemsKeepTheirHeightBasisDuringIntrinsicMeasurement(string display, bool throughContents,
        bool ordinaryWrapperInsideItem) {
        string content = SquareImage("100%");
        if (ordinaryWrapperInsideItem) content = "<span style='height:90px'>" + content + "</span>";
        string item = "<span style='height:20px;align-self:flex-start'>" + content + "</span>";
        if (throughContents) {
            item = "<span style='display:contents;height:70px'><span style='display:contents;height:90px'>"
                + item + "</span></span>";
        }
        HtmlRenderDocument rendered = Render(ShrinkToFit("",
            "<span style='display:" + display + ";height:120px;vertical-align:top'>" + item + "</span>"));

        AssertReplacedContribution(rendered, 20D, 20D);
    }

    [Theory]
    [InlineData("inline-block", "100%")]
    [InlineData("inline", "20px")]
    public void IntrinsicMeasurementRetainsRealBoxAndAbsoluteImageBoundaries(string display, string imageHeight) {
        HtmlRenderDocument rendered = Render(ShrinkToFit("height:120px",
            "<span style='display:" + display + ";height:20px;vertical-align:top'>"
            + SquareImage(imageHeight) + "</span>"));

        AssertReplacedContribution(rendered, 20D, 20D);
    }

    [Theory]
    [InlineData("content-box", 120D, 144D)]
    [InlineData("border-box", 96D, 120D)]
    public void IntrinsicReplacedWidthsDeductContainingBoxInsetsOnce(string sizing, double expectedSize,
        double expectedOuterWidth) {
        HtmlRenderDocument rendered = Render(ShrinkToFit("height:120px;padding:10px;border:2px solid red;box-sizing:" + sizing,
            "<span style='height:20px'>" + SquareImage("100%") + "</span>"));

        AssertReplacedContribution(rendered, expectedSize, expectedOuterWidth);
    }

    private static void AssertReplacedContribution(HtmlRenderDocument rendered, double expectedSize, double expectedOuterWidth) {
        HtmlRenderImage image = Assert.Single(Visuals(rendered).OfType<HtmlRenderImage>());
        Assert.Equal(expectedSize, image.Width, 6);
        Assert.Equal(expectedSize, image.Height, 6);
        Assert.Equal(expectedOuterWidth, Shape(rendered, "span#outer").Width, 6);
        Assert.Equal(expectedOuterWidth, Shape(rendered, "span#after").X, 6);
        rendered.RequireNoLoss();
    }

    private static string ShrinkToFit(string constraints, string content) =>
        "<span id='outer' style='display:inline-block;background:red;vertical-align:top;" + constraints + "'>"
        + content + "</span><span id='after' style='display:inline-block;width:10px;height:10px;background:blue;vertical-align:top'></span>";
    private static string SquareImage(string height) =>
        "<img id='image' style='height:" + height + ";width:auto;vertical-align:top' src='data:image/png;base64," + Pixel + "'>";
    private const string Pixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg==";
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
