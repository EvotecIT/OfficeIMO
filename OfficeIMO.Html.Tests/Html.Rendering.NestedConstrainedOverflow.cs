using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("height:200px")]
    [InlineData("height:80px")]
    [InlineData("max-height:80px")]
    [InlineData("min-height:200px")]
    [InlineData("aspect-ratio:3.5;min-height:0")]
    public void HtmlNestedConstrainedOverflow_UsesIndependentOverflowWithoutMovingFollowingFlow(string outerConstraint) {
        string html = FixedHeightOverflowSource("<section style='" + outerConstraint + "'>"
            + NestedConstrainedKeepContent() + "</section>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        AssertNestedConstrainedKeep(rendered);
    }

    [Theory]
    [InlineData("<article>", "</article>")]
    [InlineData("<article style='height:200px'>", "</article>")]
    public void HtmlNestedConstrainedOverflow_ImportsNestedSourcesThroughOrdinaryWrappers(string before, string after) {
        string html = FixedHeightOverflowSource("<section style='height:200px'>" + before
            + NestedConstrainedKeepContent() + after + "</section>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        AssertNestedConstrainedKeep(rendered);
    }

    [Fact]
    public void HtmlNestedConstrainedOverflow_CollapsedEntryAndLaterLineGapDoNotShiftAnIndependentKeep() {
        string html = FixedHeightOverflowSource("<section style='height:200px'>"
            + "<div style='height:60px;margin-bottom:20px'>KEEP-LEAD</div>"
            + "<div style='height:20px;margin-top:-10px'><p>KEEP-INNER</p>"
            + "<p style='break-inside:avoid;orphans:1;widows:1'>KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p></div>"
            + "<p style='padding-left:300px'>KEEP-AFTER</p></section>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "KEEP-LEAD", "KEEP-INNER", "KEEP-ONE", "KEEP-TWO", "KEEP-THREE", "KEEP-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-INNER", 1, 70D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-AFTER", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-ONE", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-TWO", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-THREE", 2, 40D);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlNestedConstrainedOverflow_AncestorEntryMoveRebuildsItsDescendantsAtTheNewPageOrigin() {
        string html = FixedHeightOverflowSource("<div style='height:60px'>KEEP-LEAD</div>"
            + "<section style='height:80px;break-inside:avoid'>"
            + "<div style='height:20px'><p>KEEP-INNER</p>"
            + "<p style='break-inside:avoid;orphans:1;widows:1'>KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p></div>"
            + "<p style='padding-left:300px'>KEEP-AFTER</p></section>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "KEEP-LEAD", "KEEP-INNER", "KEEP-ONE", "KEEP-TWO", "KEEP-THREE", "KEEP-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-INNER", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-AFTER", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-ONE", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-TWO", 2, 40D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-THREE", 2, 60D);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlNestedConstrainedOverflow_AvoidedAutoWrapperRebuildsOrdinaryChildrenAtItsFreshEntry() {
        string html = FixedHeightOverflowSource("<div style='height:60px'>KEEP-LEAD</div>"
            + "<section style='height:200px'><div style='break-inside:avoid'><p>KEEP-INNER</p>"
            + "<p style='break-inside:avoid'>KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p></div>"
            + "<p style='padding-left:300px'>KEEP-AFTER</p></section>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        // The third page is the genuine outer allocation, rather than moved text.
        Assert.Equal(3, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "KEEP-LEAD", "KEEP-INNER", "KEEP-ONE", "KEEP-TWO", "KEEP-THREE", "KEEP-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-INNER", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-ONE", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-TWO", 2, 40D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-THREE", 2, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-AFTER", 2, 80D);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlNestedConstrainedOverflow_CollapsedAuthoredMarginDoesNotReduceTheSourceEntryGap() {
        string html = FixedHeightOverflowSource("<section style='height:200px'>"
            + "<div style='height:40px;margin-bottom:20px'>KEEP-LEAD</div>"
            + "<p style='margin-top:20px;break-inside:avoid'>KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p>"
            + "<p style='padding-left:300px'>KEEP-AFTER</p></section>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "KEEP-LEAD", "KEEP-ONE", "KEEP-TWO", "KEEP-THREE", "KEEP-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-ONE", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-TWO", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-THREE", 2, 40D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-AFTER", 2, 60D);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void HtmlNestedConstrainedOverflow_SiblingKeepRebuildsProjectedChildrenAtItsInteriorCut(
        bool avoidBefore, bool clippedPrevious) {
        string previous = clippedPrevious
            ? "<p style='height:20px;overflow:hidden;break-after:avoid'>KEEP-INNER</p>"
            : "<p style='" + (avoidBefore ? "" : "break-after:avoid") + "'>KEEP-INNER</p>";
        string html = FixedHeightOverflowSource("<section style='height:200px'>"
            + "<div style='height:60px'>KEEP-LEAD</div>" + previous
            + "<p style='break-inside:avoid;" + (avoidBefore ? "break-before:avoid" : "") + "'>"
            + "KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p><p>KEEP-AFTER</p></section>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "KEEP-LEAD", "KEEP-INNER", "KEEP-ONE", "KEEP-TWO", "KEEP-THREE", "KEEP-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-INNER", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-ONE", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-TWO", 2, 40D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-THREE", 2, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-AFTER", 2, 80D);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlNestedConstrainedOverflow_TwoSiblingKeepsAdvanceThroughDistinctInteriorEntries() {
        string html = FixedHeightOverflowSource("<section style='height:400px'>"
            + "<div style='height:60px'>LEAD</div><p style='break-after:avoid'>PRE1</p>"
            + "<p style='break-inside:avoid'>ONE1<br>TWO1<br>THREE1</p>"
            + "<div style='height:60px'>BETWEEN</div><p style='break-after:avoid'>PRE2</p>"
            + "<p style='break-inside:avoid'>ONE2<br>TWO2<br>THREE2</p><p>AFTER</p></section>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(4, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "LEAD", "PRE1", "ONE1", "TWO1", "THREE1", "BETWEEN", "PRE2", "ONE2", "TWO2", "THREE2", "AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "PRE1", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "ONE1", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "BETWEEN", 2, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "PRE2", 4, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "ONE2", 4, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "TWO2", 4, 40D);
        AssertFixedHeightOverflowPlacement(rendered, "THREE2", 4, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "AFTER", 4, 80D);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("hidden")]
    [InlineData("clip")]
    public void HtmlNestedConstrainedOverflow_AtomicSiblingRetainsItsVisibleViewport(string overflow) {
        string html = FixedHeightOverflowSource("<section style='height:200px'>"
            + "<div style='height:60px'>LEAD</div><div style='height:20px'><p>X1</p><p>X2</p><p>X3</p></div>"
            + "<div style='height:20px;overflow:" + overflow + "'><p style='background:#00aa00;break-inside:avoid'>"
            + "CLIP1<br>CLIP2<br>CLIP3</p></div><p>AFTER</p></section>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(2, rendered.Pages.Count);
        AssertFixedHeightOverflowPlacement(rendered, "X1", 1, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "X2", 1, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "X3", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "CLIP1", 1, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "AFTER", 2, 0D);
        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        OfficeRasterImage second = OfficeDrawingRasterRenderer.Render(rendered.Pages[1].CreateDrawing());
        Assert.Equal(OfficeColor.FromRgb(0, 170, 0), first.GetPixel(500, 90));
        Assert.Equal(OfficeColor.White, second.GetPixel(500, 10));
    }

    private static string NestedConstrainedKeepContent() => "<div style='height:60px'>KEEP-LEAD</div>"
        + "<div style='height:20px'><p>KEEP-INNER</p>"
        + "<p style='break-inside:avoid;orphans:1;widows:1'>KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p></div>"
        + "<p style='padding-left:300px'>KEEP-AFTER</p>";

    private static void AssertNestedConstrainedKeep(HtmlRenderDocument rendered) {
        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "KEEP-LEAD", "KEEP-INNER", "KEEP-ONE", "KEEP-TWO", "KEEP-THREE", "KEEP-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-INNER", 1, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-AFTER", 1, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-ONE", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-TWO", 2, 20D);
        AssertFixedHeightOverflowPlacement(rendered, "KEEP-THREE", 2, 40D);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ConstrainedOverflowPaginationPending
            || diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
        rendered.RequireNoLoss();
    }
}
