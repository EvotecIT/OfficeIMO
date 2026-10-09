using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void TextDecorationAnonymousFlexTextPreservesAllExplicitGeometry() {
        var scene = HtmlRenderTestDriver.Render("<div style='display:flex;font:32px/40px Arial;text-decoration:underline overline line-through 5px solid red;text-underline-offset:12px;text-decoration-skip-ink:none'>Flex</div>");
        var visuals = scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).ToArray();
        HtmlRenderText text = Assert.Single(visuals.OfType<HtmlRenderText>(), t => t.Text == "Flex");
        HtmlRenderShape[] lines = DecorationLines(scene);
        Assert.Equal(3, lines.Length);
        Assert.All(lines, line => {
            Assert.Equal(5D, line.Height);
            Assert.Equal(OfficeColor.Red, line.Shape.FillColor);
            Assert.Equal(text.TextAdvanceWidth!.Value, line.Width, 6);
        });
        Assert.Equal(text.Y + 44D, Assert.Single(lines, s => s.Source!.EndsWith(":underline", StringComparison.Ordinal)).Y, 6);
        Assert.False(text.Font.IsUnderline);
        Assert.False(text.Font.IsStrikethrough);
        Assert.False(scene.HasLoss);
    }

    [Theory]
    [InlineData("<span style='display:contents;text-decoration:underline overline line-through 5px solid red'>Text</span>")]
    [InlineData("<u style='display:contents'>Text</u>")]
    [InlineData("<a href='https://example.test' style='display:contents'>Text</a>")]
    public void TextDecorationContentsDoesNotOriginateABand(string html) {
        var scene = HtmlRenderTestDriver.Render(html);
        Assert.Empty(DecorationLines(scene));
        Assert.All(scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).OfType<HtmlRenderText>(), text => {
            Assert.False(text.Font.IsUnderline);
            Assert.False(text.Font.IsStrikethrough);
        });
        Assert.DoesNotContain(scene.Diagnostics, d => d.Code is "HtmlRenderTextDecorationThicknessApproximated" or "HtmlRenderTextDecorationLineUnsupported");
    }

    [Fact]
    public void TextDecorationContentsForwardsAncestorWithoutAddingItsOwnDeclaration() {
        var scene = HtmlRenderTestDriver.Render("<span style='text-decoration:underline 4px solid red;text-decoration-skip-ink:none'><span style='display:contents;text-decoration:overline 8px solid blue'><span>Text</span></span></span>");
        HtmlRenderShape line = Assert.Single(DecorationLines(scene));
        Assert.EndsWith(":decoration:underline", line.Source);
        Assert.Equal(4D, line.Height);
        Assert.Equal(OfficeColor.Red, line.Shape.FillColor);
        Assert.False(scene.HasLoss);
    }

    [Theory]
    [InlineData("position:absolute;left:150px;top:20px")]
    [InlineData("position:fixed;left:150px;top:20px")]
    [InlineData("float:left")]
    public void TextDecorationAncestorStopsAtOutOfFlowTextButOwnBandStillPaints(string declarations) {
        const string decoration = "text-decoration:underline overline line-through 4px solid red;text-decoration-skip-ink:none";
        string html = "<div style='width:500px'><span style='" + decoration + "'>Flow<span style='" + declarations + "'>Outside</span></span></div>";
        var scene = HtmlRenderTestDriver.Render(html);
        Assert.Equal(3, DecorationLines(scene).Length);
        HtmlRenderText outside = Assert.Single(scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).OfType<HtmlRenderText>(), t => t.Text == "Outside");
        Assert.False(outside.Font.IsUnderline);
        Assert.False(outside.Font.IsStrikethrough);
        scene = HtmlRenderTestDriver.Render(html.Replace(declarations + "'", declarations + ";" + decoration + "'"));
        Assert.Equal(6, DecorationLines(scene).Length);
    }

    [Theory]
    [InlineData("auto")]
    [InlineData("all")]
    public void TextDecorationSkipInkDoesNotReportLossForLineThroughOnly(string skipInk) {
        var scene = HtmlRenderTestDriver.Render("<span style='text-decoration:line-through 4px solid red;text-decoration-skip-ink:" + skipInk + ";text-underline-position:under'>Strike</span>");
        Assert.Single(DecorationLines(scene));
        Assert.False(scene.HasLoss);
    }

    [Fact]
    public void TextDecorationNoneDescendantKeepsAncestorBandWithoutIndependentBandLoss() {
        var scene = HtmlRenderTestDriver.Render("<span style='text-decoration:underline overline line-through 4px solid red;text-decoration-skip-ink:none'><span style='text-decoration:none'>Text</span></span>");
        Assert.Equal(3, DecorationLines(scene).Length);
        Assert.False(scene.HasLoss);
    }

    [Fact]
    public void TextDecorationContinuousFootnoteRetainsTheInFlowAncestorBand() {
        var scene = HtmlRenderTestDriver.Render("<span style='text-decoration:underline 4px solid red;text-decoration-skip-ink:none'><span style='float:footnote'>Note</span></span>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous });
        Assert.Single(DecorationLines(scene));
        Assert.False(scene.HasLoss);
    }

    [Theory]
    [InlineData("margin-left:8px")]
    [InlineData("padding-right:6px")]
    [InlineData("border:2px solid blue")]
    public void TextDecorationDescendantInlineEdgesRetainSpecificContinuityLoss(string declarations) {
        var scene = HtmlRenderTestDriver.Render("<span style='text-decoration:underline 4px solid red;text-decoration-skip-ink:none'>Left<span style='" + declarations + "'>Middle</span>Right</span>");
        Assert.Contains(scene.Diagnostics, d => d.Code == "HtmlRenderTextDecorationThicknessApproximated"
            && d.Detail?.Contains("continuity across descendant inline", StringComparison.Ordinal) == true);
        Assert.True(scene.HasLoss);
        Assert.NotEmpty(DecorationLines(scene));
    }

    private static HtmlRenderShape[] DecorationLines(HtmlRenderDocument scene) =>
        scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).OfType<HtmlRenderShape>()
            .Where(s => s.Source?.Contains(":decoration:", StringComparison.Ordinal) == true).ToArray();
}
