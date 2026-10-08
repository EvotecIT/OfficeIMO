using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    // Two atomic inline boxes have a 50px min-content and 90px max-content
    // contribution. This avoids coupling sizing assertions to host font metrics.
    private const string IntrinsicWidthAtoms = "<span style='display:inline-block;width:40px;height:12px;background:navy'></span>"
        + "<span style='display:inline-block;width:50px;height:12px;background:red'></span>";

    [Theory]
    [InlineData("min-content", 300D, 50D)]
    [InlineData("max-content", 70D, 90D)]
    [InlineData("fit-content", 300D, 90D)]
    [InlineData("fit-content", 70D, 70D)]
    [InlineData("fit-content", 30D, 50D)]
    [InlineData("fit-content(70px)", 300D, 70D)]
    [InlineData("fit-content(calc(30px + 40px))", 300D, 70D)]
    public void HtmlIntrinsicWidth_BlockUsesContentContributions(string value, double parentWidth, double expected) {
        HtmlRenderDocument rendered = RenderIntrinsicWidth("width:" + value, parentWidth);

        Assert.Equal(expected, TableGeometryShape(rendered, "div#sized").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("width:min-content", "content-box", 70D)]
    [InlineData("width:min-content", "border-box", 70D)]
    [InlineData("width:max-content", "content-box", 110D)]
    [InlineData("width:max-content", "border-box", 110D)]
    [InlineData("width:fit-content(70px)", "content-box", 90D)]
    [InlineData("width:fit-content(70px)", "border-box", 70D)]
    [InlineData("width:20px;min-width:max-content;max-width:50px", "content-box", 110D)]
    [InlineData("width:20px;min-width:max-content;max-width:50px", "border-box", 110D)]
    [InlineData("width:200px;max-width:min-content", "content-box", 70D)]
    [InlineData("width:200px;max-width:min-content", "border-box", 70D)]
    public void HtmlIntrinsicWidth_ConstraintsAndBoxEdgesUseTheCorrectCoordinates(string declarations, string sizing, double expected) {
        HtmlRenderDocument rendered = RenderIntrinsicWidth(declarations + ";box-sizing:" + sizing + ";padding:8px;border:2px solid black", 300D);

        Assert.Equal(expected, TableGeometryShape(rendered, "div#sized").Width, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlIntrinsicWidth_ResolvedWidthPrecedesCenteredAutoMargins() {
        HtmlRenderDocument rendered = RenderIntrinsicWidth("width:fit-content;margin:0 auto;padding:10px", 300D);
        HtmlRenderShape shape = TableGeometryShape(rendered, "div#sized");

        Assert.Equal(110D, shape.Width, 3);
        Assert.Equal(95D, shape.X, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlIntrinsicWidth_InlineBlockMaxContentCanExceedItsContainingWidth() {
        string html = TableGeometrySource("<div style='width:70px'><span id='sized' style='display:inline-block;width:max-content;background:lime'>"
            + IntrinsicWidthAtoms + "</span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(90D, TableGeometryShape(rendered, "span#sized").Width, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlIntrinsicWidth_NestedAuthoredWidthParticipatesInParentContribution() {
        string html = TableGeometrySource("<div id='sized' style='width:min-content;background:lime'>"
            + "<div style='width:max-content'>" + IntrinsicWidthAtoms + "</div></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(90D, TableGeometryShape(rendered, "div#sized").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("width:max-content", "width:min-content", 50D)]
    [InlineData("width:min-content", "width:max-content", 90D)]
    [InlineData("width:max-content", "width:100%;max-width:60px", 60D)]
    [InlineData("width:max-content", "width:25%;min-width:80px", 90D)]
    [InlineData("width:max-content", "min-width:100%", 90D)]
    [InlineData("width:max-content", "width:fit-content(70px)", 70D)]
    public void HtmlIntrinsicWidth_NestedConstraintsAndCyclicPercentagesUseIndefiniteContributions(string outer, string inner, double expected) {
        string html = TableGeometrySource("<div id='sized' style='background:lime;" + outer + "'><div style='" + inner + "'>"
            + IntrinsicWidthAtoms + "</div></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(expected, TableGeometryShape(rendered, "div#sized").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("min-inline-size:max-content;width:20px", 90D)]
    [InlineData("max-inline-size:min-content;width:200px", 50D)]
    [InlineData("inline-size:fit-content;--limit:70px;max-inline-size:var(--limit)", 70D)]
    public void HtmlIntrinsicWidth_LogicalDimensionsShareTheHorizontalSizingOwner(string declarations, double expected) {
        HtmlRenderDocument rendered = RenderIntrinsicWidth(declarations, 300D);

        Assert.Equal(expected, TableGeometryShape(rendered, "div#sized").Width, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlIntrinsicWidth_NestedInlineBlockIsAtomicInParentMinimum() {
        string html = TableGeometrySource("<div id='sized' style='background:lime;width:min-content'>"
            + "<span style='display:inline-block;width:max-content;padding:5px'>" + IntrinsicWidthAtoms + "</span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());

        Assert.Equal(100D, TableGeometryShape(rendered, "div#sized").Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("normal")]
    [InlineData("nowrap")]
    public void HtmlIntrinsicWidth_StyledTextMeasurementAgreesWithPaintedUnbrokenText(string whitespace) {
        string html = TableGeometrySource("<div style='width:300px'><div id='sized' style='width:max-content;background:lime;"
            + "font-size:19px;font-weight:bold;letter-spacing:2px;word-spacing:3px;white-space:" + whitespace
            + "'>Longword short</div></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlRenderShape box = TableGeometryShape(rendered, "div#sized");
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();

        Assert.NotEmpty(text);
        Assert.Equal(text.Max(item => item.X + item.Width) - box.X, box.Width, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("position:absolute", "")]
    [InlineData("float:left", "")]
    [InlineData("display:flex", "")]
    [InlineData("writing-mode:vertical-rl", "")]
    [InlineData("", "<table><tr><td>Cell</td></tr></table>")]
    [InlineData("", "<div style='display:grid'>Grid</div>")]
    [InlineData("width:fit-content(50%)", "")]
    public void HtmlIntrinsicWidth_UnqualifiedContextsRetainPropertySpecificLoss(string declarations, string content) {
        string html = TableGeometrySource("<div id='sized' style='width:max-content;" + declarations + "'>"
            + (content.Length == 0 ? IntrinsicWidthAtoms : content) + "</div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, TableGeometryOptions());
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics,
            item => item.Code == HtmlRenderDiagnosticCodes.IntrinsicSizeUnsupported && item.Source == "div#sized");

        Assert.Contains("width=", diagnostic.Detail);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    private static HtmlRenderDocument RenderIntrinsicWidth(string declarations, double parentWidth) {
        string html = TableGeometrySource("<div style='width:" + parentWidth.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "px'><div id='sized' style='background:lime;" + declarations + "'>" + IntrinsicWidthAtoms + "</div></div>");
        return HtmlRenderTestDriver.Render(html, TableGeometryOptions());
    }
}
