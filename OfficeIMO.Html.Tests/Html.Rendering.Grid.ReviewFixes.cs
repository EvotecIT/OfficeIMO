using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("inline", "grid-template-columns", "40px 60px", "garbage", "grid-column:2", 40D, 0D, 60D)]
    [InlineData("rule", "grid-template-columns", "40px 60px", "garbage", "grid-column:2", 40D, 0D, 60D)]
    [InlineData("rules", "grid-template-columns", "40px 60px", "garbage!important", "grid-column:2", 40D, 0D, 60D)]
    [InlineData("inline", "grid-template-rows", "20px 30px", "garbage", "grid-row:2", 0D, 20D, 40D)]
    [InlineData("rule", "grid-template-rows", "20px 30px", "garbage", "grid-row:2", 0D, 20D, 40D)]
    [InlineData("rules", "grid-template-rows", "20px 30px", "garbage!important", "grid-row:2", 0D, 20D, 40D)]
    [InlineData("inline", "grid-template-areas", "\"first second\"", "\"first second\" \"first\"", "grid-area:second", 40D, 0D, 60D)]
    [InlineData("rule", "grid-template-areas", "\"first second\"", "\"first second\" \"first\"", "grid-area:second", 40D, 0D, 60D)]
    [InlineData("rules", "grid-template-areas", "\"first second\"", "\"first second\" \"first\"!important", "grid-area:second", 40D, 0D, 60D)]
    [InlineData("rule", "grid-template-areas", "\"first second\"", "\"first second\" \"second first\"", "grid-area:second", 40D, 0D, 60D)]
    [InlineData("rule", "grid-template-columns", "40px 60px", "repeat(0,20px)", "grid-column:2", 40D, 0D, 60D)]
    [InlineData("rule", "grid-template-columns", "40px 60px", "repeat(auto-fit,repeat(auto-fit,1px))", "grid-column:2", 40D, 0D, 60D)]
    [InlineData("rule", "grid-template-columns", "40px 60px", "subgrid garbage", "grid-column:2", 40D, 0D, 60D)]
    [InlineData("rule", "grid-template-columns", "40px 60px", "minmax(min(garbage,20px),1fr)", "grid-column:2", 40D, 0D, 60D)]
    public void HtmlGrid_InvalidTemplatePreservesEarlierValidDeclaration(string scope, string property, string valid,
        string invalid, string placement, double expectedX, double expectedY, double expectedWidth) {
        string declarations = property + ":" + valid + ";" + property + ":" + invalid;
        string css = scope == "rule" ? "<style>#grid{" + declarations + "}</style>"
            : scope == "rules" ? "<style>#grid{" + property + ":" + valid + "}#grid{" + property + ":" + invalid + "}</style>" : "";
        string html = css + "<div id='grid' style='display:grid;width:100px;"
            + (property != "grid-template-columns" || scope == "inline" ? "grid-template-columns:40px 60px;" : "")
            + (scope == "inline" ? declarations : "") + "'><span id='last' style='background:red;" + placement + "'>B</span></div>";
        HtmlRenderShape item = FindGridShape(RenderGrid(html, 120D), "span#last");
        Assert.Equal(expectedX, item.X, 3);
        Assert.Equal(expectedY, item.Y, 3);
        Assert.Equal(expectedWidth, item.Width, 3);
        if (property == "grid-template-rows") Assert.Equal(30D, item.Height, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlGrid_IntrinsicMeasurementDefersAbsolutePercentageResolution(bool nested) {
        string children = "<span style='width:100px'>A</span><div style='display:contents'>"
            + "<span id='absolute' style='position:absolute;width:50%;min-width:25%;height:10px;background:red'>B</span></div>";
        if (nested) children = "<div style='display:grid;grid-column:1;grid-template-columns:subgrid'>" + children + "</div>";
        string html = "<div style='display:grid;width:400px;grid-template-columns:100px 300px'>"
            + "<div style='position:relative;display:grid;grid-column:1;grid-template-columns:subgrid'>" + children + "</div></div>";
        HtmlRenderShape absolute = FindGridShape(RenderGrid(html, 600D), "span#absolute");
        Assert.Equal(50D, absolute.Width, 3);
        Assert.Equal(10D, absolute.Height, 3);
    }

    [Theory]
    [InlineData("inline-grid")]
    [InlineData("inline-flex")]
    public void HtmlGrid_TrueInlineIntrinsicMeasurementDefersAbsolutePercentageResolution(string display) {
        string html = "<p style='width:400px;margin:0'>"
            + $"<span style='display:{display};position:relative;grid-template-columns:100px'>"
            + "<span style='width:100px'>A</span><span id='absolute' style='position:absolute;width:50%;height:10px;background:red'>B</span>"
            + "</span></p>";
        HtmlRenderShape absolute = FindGridShape(RenderGrid(html, 600D), "span#absolute");
        Assert.Equal(50D, absolute.Width, 3);
        Assert.Equal(10D, absolute.Height, 3);
    }

    [Theory]
    [InlineData(0D, "100px", false, 80D, 260D)]
    [InlineData(20D, "100px", false, 80D, 280D)]
    [InlineData(100D, "20px", false, 80D, 280D)]
    [InlineData(20D, "normal", false, 80D, 200D)]
    [InlineData(20D, "100px", true, 80D, 220D)]
    public void HtmlGrid_SubgridIntrinsicWidthsAccountForDifferentAndInheritedGaps(double parentGap, string childGap,
        bool nested, double childWidth, double expectedTailX) {
        string children = $"<span id='gap-a' style='width:{childWidth}px;background:red'>A</span>"
            + $"<span id='gap-b' style='width:{childWidth}px;background:blue'>B</span>";
        if (nested) children = "<div style='display:grid;grid-column:1 / 3;grid-template-columns:subgrid;column-gap:40px'>" + children + "</div>";
        string html = $"<div style='display:grid;width:600px;grid-template-columns:auto auto minmax(0,1fr);column-gap:{parentGap}px'>"
            + "<div style='display:grid;grid-column:1 / 3;grid-template-columns:subgrid;"
            + (childGap == "normal" ? "" : "column-gap:" + childGap + ";") + "'>" + children + "</div>"
            + "<span id='tail' style='grid-column:3;background:green'>C</span></div>";
        HtmlRenderDocument rendered = RenderGrid(html, 620D);
        HtmlRenderShape first = FindGridShape(rendered, "span#gap-a");
        HtmlRenderShape second = FindGridShape(rendered, "span#gap-b");
        Assert.Equal(childWidth, first.Width, 3);
        Assert.Equal(childWidth, second.Width, 3);
        Assert.Equal(expectedTailX, FindGridShape(rendered, "span#tail").X, 3);
        Assert.True(second.X >= first.X + first.Width);
    }

    [Fact]
    public void HtmlGrid_PreservedTemplateRetainsSubgridAndModernLengthMath() {
        const string html = "<style>#grid{grid-template-columns:repeat(auto-fit,minmax(min(100%,148px),1fr))}"
            + "#subgrid{grid-template-columns:subgrid [start] [end]}</style>"
            + "<div id='grid' style='display:grid;width:300px'><div id='subgrid' style='display:grid;grid-column:1 / -1'>"
            + "<span id='modern' style='grid-column:2;background:red'>A</span></div></div>";
        HtmlRenderDocument rendered = RenderGrid(html, 320D);
        HtmlRenderShape item = FindGridShape(rendered, "span#modern");
        Assert.Equal(150D, item.X, 3);
        Assert.Equal(150D, item.Width, 3);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GridValueUnsupported);
    }

    [Fact]
    public void HtmlGrid_SpanningSubgridContributionDoesNotAddAnInteriorGapToOuterEdges() {
        const string html = "<div style='display:grid;width:600px;grid-template-columns:auto auto minmax(0,1fr)'>"
            + "<div style='display:grid;grid-column:1 / 3;grid-template-columns:subgrid;column-gap:100px'>"
            + "<span id='spanning' style='grid-column:1 / 3;width:200px;background:red'>A</span></div>"
            + "<span id='tail' style='grid-column:3;background:green'>B</span></div>";
        HtmlRenderDocument rendered = RenderGrid(html, 620D);
        Assert.Equal(200D, FindGridShape(rendered, "span#spanning").Width, 3);
        Assert.Equal(200D, FindGridShape(rendered, "span#tail").X, 3);
    }
}
