using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private const string MixedTableMathRatio = "((50% + 10px) / (25% + 10px))";

    [Theory]
    [InlineData("cell", "25%")]
    [InlineData("css-cell", "25%")]
    [InlineData("column", "25%")]
    [InlineData("cell", "10px")]
    public void HtmlTableMathRatios_MixedNumericFactorsUseAutoWidthsAcrossSharedInternalOwners(string owner, string dimension) {
        string width = "calc(" + dimension + " * " + MixedTableMathRatio + ")";
        bool css = owner == "css-cell";
        string html = TableGeometrySource(css
            ? "<div style='display:table;width:240px;table-layout:fixed;margin:0;border-spacing:0'><div style='display:table-row'>"
                + "<div id='summary' style='display:table-cell;padding:0;background:lime;width:" + width + "'>Summary</div>"
                + "<div id='detail' style='display:table-cell;padding:0;background:lime'>Detail</div></div></div>"
            : "<table style='width:240px;table-layout:fixed;margin:0;border-spacing:0'>"
                + (owner == "column" ? "<colgroup><col style='width:" + width + "'><col></colgroup>" : "")
                + "<tr><td id='summary' style='padding:0;background:lime;" + (owner == "cell" ? "width:" + width : "") + "'>Summary</td>"
                + "<td id='detail' style='padding:0;background:lime'>Detail</td></tr></table>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(120D, TableGeometryShape(rendered, (css ? "div" : "td") + "#summary").Width, 3);
        Assert.Equal(120D, TableGeometryShape(rendered, (css ? "div" : "td") + "#detail").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("cell")]
    [InlineData("css-cell")]
    [InlineData("row")]
    public void HtmlTableMathRatios_MixedNumericFactorsUseAutoBeforePercentageRowAllocation(string owner) {
        string height = "height:calc(25% * " + MixedTableMathRatio + ")";
        bool css = owner == "css-cell";
        string html = TableGeometrySource((css
            ? "<div style='display:table;height:200px;width:240px;margin:0;border-spacing:0'><div style='display:table-row'>"
                + "<div id='summary' style='display:table-cell;padding:0;background:lime;" + height + "'>Summary</div></div>"
                + "<div style='display:table-row'><div id='detail' style='display:table-cell;height:50%;padding:0;background:lime'>Detail</div></div></div>"
            : "<table style='height:200px;width:240px;margin:0;border-spacing:0'><tr style='" + (owner == "row" ? height : "") + "'>"
                + "<td id='summary' style='padding:0;background:lime;" + (owner == "cell" ? height : "") + "'>Summary</td></tr>"
                + "<tr><td id='detail' style='height:50%;padding:0;background:lime'>Detail</td></tr></table>")
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(100D, TableGeometryShape(rendered, (css ? "div" : "td") + "#summary").Height, 3);
        Assert.Equal(100D, TableGeometryShape(rendered, (css ? "div" : "td") + "#detail").Height, 3);
        Assert.Equal(200D, TableGeometryShape(rendered, "div#next").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("calc(25% * (30px / 10px))")]
    [InlineData("calc(25% * (75% / 25%))")]
    [InlineData("calc(25% * ((75% + 0px) / (25% + 0px)))")]
    [InlineData("calc(25% * ((75% + 10px - 10px) / (25% + 10px - 10px)))")]
    [InlineData("calc(75% + (25% + 10px) * 0)")]
    public void HtmlTableMathRatios_PureCancelledAndZeroFactorsPreserveComputedPercentageDimensions(string expression) {
        string html = TableGeometrySource("<table style='height:200px;width:240px;table-layout:fixed;margin:0;border-spacing:0'><tr>"
            + "<td id='summary' style='padding:0;background:lime;width:" + expression + ";height:" + expression + "'>Summary</td>"
            + "<td id='detail' style='padding:0;background:lime'>Detail</td></tr>"
            + "<tr><td id='tail' colspan='2' style='padding:0;background:lime'>Tail</td></tr></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(180D, TableGeometryShape(rendered, "td#summary").Width, 3);
        Assert.Equal(60D, TableGeometryShape(rendered, "td#detail").Width, 3);
        Assert.Equal(150D, TableGeometryShape(rendered, "td#summary").Height, 3);
        Assert.Equal(50D, TableGeometryShape(rendered, "td#tail").Height, 3);
        Assert.Equal(200D, TableGeometryShape(rendered, "div#next").Y, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlTableMathRatios_ZeroMultiplierDoesNotEraseAnUnresolvedMixedQuotient() {
        string expression = "calc(75% + 25% * " + MixedTableMathRatio + " * 0)";
        string html = TableGeometrySource("<table style='height:200px;width:240px;table-layout:fixed;margin:0;border-spacing:0'><tr>"
            + "<td id='summary' style='padding:0;background:lime;width:" + expression + ";height:" + expression + "'>Summary</td>"
            + "<td id='detail' style='padding:0;background:lime'>Detail</td></tr>"
            + "<tr><td id='tail' colspan='2' style='padding:0;background:lime'>Tail</td></tr></table>"
            + "<div id='next' style='height:20px;background:red'>Next</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(120D, TableGeometryShape(rendered, "td#summary").Width, 3);
        Assert.Equal(120D, TableGeometryShape(rendered, "td#detail").Width, 3);
        Assert.Equal(100D, TableGeometryShape(rendered, "td#summary").Height, 3);
        Assert.Equal(100D, TableGeometryShape(rendered, "td#tail").Height, 3);
        Assert.Equal(200D, TableGeometryShape(rendered, "div#next").Y, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlTableMathRatios_OrdinaryBlocksRetainTheirLayoutDependentNumericFactors() {
        string html = TableGeometrySource("<div id='ordinary' style='width:calc(25% * " + MixedTableMathRatio + ");background:red'>Ordinary</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(290.625D, TableGeometryShape(rendered, "div#ordinary").Width, 3);
        rendered.RequireNoLoss();
    }
}
