using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(null, 20D, 40D)]
    [InlineData(20, 20D, 40D)]
    [InlineData(25, 25D, 45D)]
    public void HtmlNestedConstrainedOverflow_FirstContentDeferralAdvancesTheAuthoredBoxBeforeNegativeFollowingFlow(
        int? firstParagraphHeight, double secondTop, double thirdTop) {
        string firstParagraph = "<p" + (firstParagraphHeight.HasValue
            ? " style='height:" + firstParagraphHeight.Value + "px'" : "") + ">ENTRY-ONE</p>";
        string html = FixedHeightOverflowSource("<div style='height:1081px'>ENTRY-LEAD</div>"
            + "<div style='height:25px;margin-bottom:-45px;background:#00aa00'>"
            + firstParagraph + "<p>ENTRY-TWO</p><p>ENTRY-THREE</p></div><p>ENTRY-AFTER</p>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        Assert.Equal(12, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered,
            "ENTRY-LEAD", "ENTRY-ONE", "ENTRY-TWO", "ENTRY-THREE", "ENTRY-AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "ENTRY-ONE", 12, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "ENTRY-TWO", 12, secondTop);
        AssertFixedHeightOverflowPlacement(rendered, "ENTRY-THREE", 12, thirdTop);
        // The first line's 19px deferral moves the whole 25px box. Its signed
        // -20px following advance therefore begins at 1100px, not 1081px.
        AssertFixedHeightOverflowPlacement(rendered, "ENTRY-AFTER", 11, 80D);
        OfficeRasterImage preceding = OfficeDrawingRasterRenderer.Render(rendered.Pages[10].CreateDrawing());
        Assert.Equal(OfficeColor.White, preceding.GetPixel(500, 90));
        rendered.RequireNoLoss();
    }

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
    [InlineData("right", 0)]
    [InlineData("left", 1)]
    public void HtmlNestedConstrainedOverflow_PreservesAlreadyAppliedLeadingSideControl(
        string leadingSide, int leadingBlankPages) {
        string html = FixedHeightOverflowSource("<section style='height:200px'>"
            + NestedConstrainedKeepContent(leadingSide) + "</section>");

        foreach (HtmlRenderIntentProfile profile in new[] {
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, FixedHeightOverflowOptions())).Document;

            // Preserve the canonical outer paginator's first-page side policy.
            // Its leading-left blank page is an established source contract;
            // native producers can choose different initial-page parity.
            Assert.Equal(2 + leadingBlankPages, rendered.Pages.Count);
            AssertDescendantSpannerSourceOnce(rendered,
                "KEEP-LEAD", "KEEP-INNER", "KEEP-ONE", "KEEP-TWO", "KEEP-THREE", "KEEP-AFTER");
            AssertFixedHeightOverflowPlacement(rendered, "KEEP-LEAD", leadingBlankPages + 1, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "KEEP-INNER", leadingBlankPages + 1, 60D);
            AssertFixedHeightOverflowPlacement(rendered, "KEEP-AFTER", leadingBlankPages + 1, 80D);
            AssertFixedHeightOverflowPlacement(rendered, "KEEP-ONE", leadingBlankPages + 2, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "KEEP-TWO", leadingBlankPages + 2, 20D);
            AssertFixedHeightOverflowPlacement(rendered, "KEEP-THREE", leadingBlankPages + 2, 40D);
            if (leadingBlankPages > 0) Assert.Empty(rendered.Pages[0].Visuals.OfType<HtmlRenderText>());
            Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
                diagnostic.Code == HtmlRenderDiagnosticCodes.ConstrainedOverflowPaginationPending);
            rendered.RequireNoLoss();
        }
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
    public void HtmlNestedConstrainedOverflow_InteriorMarginAtThePageBoundaryKeepsItsSharedSourceDiscard() {
        string html = FixedHeightOverflowSource("<section style='height:200px'>"
            + "<div style='height:60px'>KEEP-LEAD</div><div style='height:20px'>"
            + "<p>KEEP-INNER</p><p style='margin-top:20px;break-inside:avoid'>"
            + "KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p></div>"
            + "<p style='padding-left:300px'>KEEP-AFTER</p></section>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, FixedHeightOverflowOptions());

        AssertNestedConstrainedKeep(rendered);
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

    [Theory]
    [InlineData("inline", false)]
    [InlineData("stylesheet", false)]
    [InlineData("nested", false)]
    [InlineData("inherit", false)]
    [InlineData("stylesheet", true)]
    public void HtmlNestedConstrainedOverflow_TransparentWrappersPreserveTheirResolvedSiblingKeeps(
        string route, bool avoidBefore) {
        string kept = "<p style='" + (avoidBefore ? "" : "break-after:avoid") + "'>KEEP-INNER</p>"
            + "<p style='break-inside:avoid;" + (avoidBefore ? "break-before:avoid" : "") + "'>"
            + "KEEP-ONE<br>KEEP-TWO<br>KEEP-THREE</p>";
        string wrapped = route switch {
            "stylesheet" => "<style>.transparent{display:contents}</style><div class='transparent'>" + kept + "</div>",
            "nested" => "<div style='display:contents'><div style='display:contents'>" + kept + "</div></div>",
            "inherit" => "<div style='display:contents'><div style='display:inherit'>" + kept + "</div></div>",
            _ => "<div style='display:contents'>" + kept + "</div>"
        };
        string html = FixedHeightOverflowSource("<section style='height:200px'>"
            + "<div style='height:60px'>KEEP-LEAD</div>" + wrapped
            + "<p style='padding-left:300px'>KEEP-AFTER</p></section>");
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

    private static string NestedConstrainedKeepContent(string? leadingSide = null) => "<div style='height:60px"
        + (leadingSide == null ? "" : ";break-before:" + leadingSide) + "'>KEEP-LEAD</div>"
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
