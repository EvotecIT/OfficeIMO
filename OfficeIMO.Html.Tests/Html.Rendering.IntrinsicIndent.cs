using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("min-content", "30px", 30D)]
    [InlineData("max-content", "30px", 30D)]
    [InlineData("max-content", "-10px", -10D)]
    [InlineData("max-content", "50%", 0D)]
    [InlineData("max-content", "calc(50% + 7px)", 7D)]
    public void HtmlIntrinsicIndent_ParagraphOwnsDecoratedInlineContributions(
        string sizing, string indent, double contribution) {
        var options = TableIntrinsicOptions();
        double textWidth = IntrinsicIndentTextWidth(options, "AAAA");
        string html = "<div id='target' style='width:" + sizing + ";text-indent:" + indent
            + ";background:lime'><span style='text-indent:0;padding:0 10px'>AAAA</span></div>";
        HtmlRenderDocument rendered = RenderTableIntrinsic(html, options);

        Assert.Equal(textWidth + 20D + contribution, TableGeometryShape(rendered, "div#target").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("cell")]
    [InlineData("caption")]
    public void HtmlIntrinsicIndent_TableUsesCellAndCaptionParagraphs(string owner) {
        var options = TableIntrinsicOptions();
        string content = "<span style='text-indent:0;padding:0 10px'>AAAA</span>";
        string html = owner == "cell"
            ? "<table id='target'><tr><td id='cell' style='text-indent:30px'>" + content + "</td></tr></table>"
            : "<table id='target'><caption style='text-indent:30px'>" + content
                + "</caption><tr><td>A</td></tr></table>";
        HtmlRenderDocument rendered = RenderTableIntrinsic(html, options);

        Assert.Equal(IntrinsicIndentTextWidth(options, "AAAA") + 50D,
            TableIntrinsicGroup(rendered, "table#target").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("min-content", "A AAAAAAAA", "")]
    [InlineData("max-content", "A<br>AAAAAAAA", "")]
    [InlineData("max-content", "A\nAAAAAAAA", "white-space:pre")]
    public void HtmlIntrinsicIndent_FollowingSoftAndForcedLinesUseTheirFullWidth(
        string sizing, string content, string whitespace) {
        var options = TableIntrinsicOptions();
        HtmlRenderDocument rendered = RenderTableIntrinsic("<div id='target' style='width:" + sizing
            + ";text-indent:30px;background:lime;" + whitespace + "'>" + content + "</div>", options);

        Assert.Equal(Math.Max(30D + IntrinsicIndentTextWidth(options, "A"),
            IntrinsicIndentTextWidth(options, "AAAAAAAA")), TableGeometryShape(rendered, "div#target").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("min-content")]
    [InlineData("max-content")]
    public void HtmlIntrinsicIndent_TabsRetainContentRelativeStops(string sizing) {
        var options = TableIntrinsicOptions();
        HtmlRenderDocument rendered = RenderTableIntrinsic("<div id='target' style='width:" + sizing
            + ";text-indent:30px;white-space:pre;tab-size:40px;background:lime'>A\tA</div>", options);

        Assert.Equal(30D + 40D + IntrinsicIndentTextWidth(options, "A"),
            TableGeometryShape(rendered, "div#target").Width, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlIntrinsicIndent_NestedBlockExcludesAnEmptyOuterFirstLine() {
        var options = TableIntrinsicOptions();
        HtmlRenderDocument rendered = RenderTableIntrinsic(
            "<div id='target' style='width:max-content;text-indent:100px;background:lime'>"
            + "<div style='text-indent:10px'>AAAA</div></div>", options);

        Assert.Equal(10D + IntrinsicIndentTextWidth(options, "AAAA"),
            TableGeometryShape(rendered, "div#target").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlIntrinsicIndent_AnonymousTextAfterBlockDoesNotRestartTheOuterIndent(bool wrapped) {
        var options = TableIntrinsicOptions();
        const string content = "A<div style='text-indent:10px'>AAAA</div>AAAAAAAA";
        HtmlRenderDocument rendered = RenderTableIntrinsic(
            "<div id='target' style='width:max-content;text-indent:30px;background:lime'>"
            + (wrapped ? "<span style='text-indent:80px'>" + content + "</span>" : content) + "</div>", options);

        Assert.Equal(IntrinsicIndentTextWidth(options, "AAAAAAAA"),
            TableGeometryShape(rendered, "div#target").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("inline-flex")]
    [InlineData("inline-grid")]
    public void HtmlIntrinsicIndent_FlexAndGridAggregateTheirItemParagraphsOnce(string display) {
        var options = TableIntrinsicOptions();
        HtmlRenderDocument rendered = RenderTableIntrinsic("<div id='target' style='display:" + display
            + ";grid-template-columns:max-content;text-indent:30px;background:lime'><span>AAAA</span></div>", options);

        Assert.Equal(30D + IntrinsicIndentTextWidth(options, "AAAA"),
            TableGeometryShape(rendered, "div#target").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlIntrinsicIndent_AtomicParagraphAndSurroundingLineInheritIndependently(bool generated) {
        var options = TableIntrinsicOptions();
        string before = generated ? "<style>#target::before{display:inline-block;content:'AAAA'}</style>" : "";
        string content = generated ? "" : "<span style='display:inline-block'>AAAA</span>";
        HtmlRenderDocument rendered = RenderTableIntrinsic(
            before + "<div id='target' style='width:max-content;text-indent:30px;background:lime'>"
            + content + "</div>", options);

        Assert.Equal(60D + IntrinsicIndentTextWidth(options, "AAAA"),
            TableGeometryShape(rendered, "div#target").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("2em")]
    [InlineData("2ch")]
    public void HtmlIntrinsicIndent_InheritedLengthsRetainTheirDeclaringFont(string indent) {
        var options = TableIntrinsicOptions();
        double contribution = 20D;
        if (indent == "2ch") {
            Assert.True(options.Fonts.TryMeasureText("0", 10D, "Pinned", OfficeFontStyle.Regular, out double zero));
            contribution = 2D * zero;
        }
        HtmlRenderDocument rendered = RenderTableIntrinsic(
            "<div style='font-size:10px;text-indent:" + indent + "'>"
            + "<div id='target' style='font-size:20px;width:max-content;background:lime'>AAAA</div></div>", options);

        Assert.Equal(contribution + IntrinsicIndentTextWidth(options, "AAAA"),
            TableGeometryShape(rendered, "div#target").Width, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlIntrinsicIndent_ManualHyphenationConsumesIndentOnlyOnTheInitialSegment() {
        var options = TableIntrinsicOptions();
        HtmlRenderDocument rendered = RenderTableIntrinsic(
            "<div id='target' style='width:min-content;text-indent:30px;hyphens:manual;background:lime'>"
            + "AA&shy;AAAA</div>", options);

        Assert.Equal(Math.Max(30D + IntrinsicIndentTextWidth(options, "AA-"),
            IntrinsicIndentTextWidth(options, "AAAA")), TableGeometryShape(rendered, "div#target").Width, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlIntrinsicIndent_SignedMarginsRemainSeparateFromTheParagraphIndent() {
        var options = TableIntrinsicOptions();
        HtmlRenderDocument rendered = RenderTableIntrinsic(
            "<div id='target' style='width:max-content;text-indent:30px;background:lime'>"
            + "<span style='margin:0 -5px;padding:0 10px'>AAAA</span></div>", options);

        Assert.Equal(40D + IntrinsicIndentTextWidth(options, "AAAA"),
            TableGeometryShape(rendered, "div#target").Width, 3);
        rendered.RequireNoLoss();
    }

    private static double IntrinsicIndentTextWidth(HtmlRenderOptions options, string text) {
        Assert.True(options.Fonts.TryMeasureText(text, 20D, "Pinned", OfficeFontStyle.Regular, out double width));
        return width;
    }
}
