using AngleSharp.Html.Parser;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlGapShorthandTests {
    [Theory]
    [InlineData("gap:0 6px", "0", "6px")]
    [InlineData("gap:4px", "4px", "4px")]
    [InlineData("column-gap:3px;gap:0 6px", "0", "6px")]
    [InlineData("gap:0 6px;column-gap:3px", "0", "3px")]
    [InlineData("column-gap:3px!important;gap:0 6px", "0", "3px")]
    [InlineData("gap:0 6px!important;column-gap:3px", "0", "6px")]
    [InlineData("--space:0 6px;column-gap:3px;gap:var(--space)", "0", "6px")]
    [InlineData("--space:0 6px;gap:var(--space);column-gap:3px", "0", "3px")]
    [InlineData("column-gap:3px!important;gap:var(--missing)", "", "3px")]
    [InlineData("gap:4px;gap:var(--missing)", "", "")]
    [InlineData("gap:4px;gap:-2px 6px", "4px", "4px")]
    [InlineData("gap:4px;gap:1px 2px 3px", "4px", "4px")]
    [InlineData("gap:4px;row-gap:-2px", "4px", "4px")]
    [InlineData("gap:inherit", "2px", "5px")]
    [InlineData("gap:4px;gap:initial", "", "")]
    [InlineData("gap:3px 7px;row-gap:initial", "", "7px")]
    [InlineData("gap:3px 7px;column-gap:initial", "3px", "")]
    [InlineData("gap:3px 7px;row-gap:var(--missing)", "", "7px")]
    public void GapShorthandKeepsRowColumnOrderAndCascade(string declaration, string row, string column) {
        foreach (bool inline in new[] { true, false }) {
            string html = (inline ? "" : "<style>#target{" + declaration + "}</style>")
                + "<div style='gap:2px 5px'><span id='target'"
                + (inline ? " style='" + declaration + "'" : "") + ">Text</span></div>";
            var document = new HtmlParser().ParseDocument(html);
            HtmlComputedStyle computed = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];
            Assert.Equal(row, computed.GetValue("row-gap"));
            Assert.Equal(column, computed.GetValue("column-gap"));
        }
    }

    [Fact]
    public void GapShorthandRevertLayerRestoresBothEarlierComponents() {
        var document = new HtmlParser().ParseDocument("<style>@layer base, override;"
            + "@layer base{#target{gap:2px 5px}}"
            + "@layer override{#target{gap:8px 9px;gap:revert-layer}}</style><p id='target'>Text</p>");
        HtmlComputedStyle computed = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];
        Assert.Equal("2px", computed.GetValue("row-gap"));
        Assert.Equal("5px", computed.GetValue("column-gap"));
    }
}

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("flex;flex-wrap:wrap", "gap:3px 7px", 3D, 7D)]
    [InlineData("grid;grid-template-columns:40px 40px", "gap:3px 7px", 3D, 7D)]
    [InlineData("flex;flex-wrap:wrap", "gap:3px 7px;row-gap:initial", 0D, 7D)]
    [InlineData("grid;grid-template-columns:40px 40px", "gap:3px 7px;row-gap:initial", 0D, 7D)]
    [InlineData("flex;flex-wrap:wrap", "gap:3px 7px;column-gap:initial", 3D, 0D)]
    [InlineData("grid;grid-template-columns:40px 40px", "gap:3px 7px;column-gap:initial", 3D, 0D)]
    [InlineData("flex;flex-wrap:wrap", "gap:3px 7px;row-gap:var(--missing)", 0D, 7D)]
    [InlineData("grid;grid-template-columns:40px 40px", "gap:3px 7px;row-gap:var(--missing)", 0D, 7D)]
    public void HtmlLayoutGap_UsesColumnSpacingAndRowSpacingOnTheCorrectAxes(string display, string declaration, double rowGap, double columnGap) {
        string html = "<div style='display:" + display + ";width:100px;" + declaration + "'>"
            + "<div id='first' style='width:40px;height:20px;background:red'></div>"
            + "<div id='second' style='width:40px;height:20px;background:blue'></div>"
            + "<div id='third' style='width:40px;height:20px;background:green'></div></div>";
        HtmlRenderDocument rendered = RenderFlex(html, 120D);
        HtmlRenderShape first = FindFlexShape(rendered, "div#first");
        HtmlRenderShape second = FindFlexShape(rendered, "div#second");
        HtmlRenderShape third = FindFlexShape(rendered, "div#third");
        Assert.Equal(first.X + first.Width + columnGap, second.X, 3);
        Assert.Equal(first.Y, second.Y, 3);
        Assert.Equal(first.Y + first.Height + rowGap, third.Y, 3);
        Assert.Equal(first.X, third.X, 3);
    }
}
