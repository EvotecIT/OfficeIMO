using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("unforced", 1, 1, 20D, 1, 40D, 1, 20D)]
    [InlineData("before-second", 2, 2, 0D, 2, 20D, 1, 20D)]
    [InlineData("before-third", 2, 1, 20D, 2, 0D, 1, 20D)]
    [InlineData("after-first", 2, 2, 0D, 2, 20D, 2, 0D)]
    [InlineData("after-second", 2, 1, 20D, 2, 0D, 2, 0D)]
    [InlineData("adjacent", 2, 2, 0D, 2, 20D, 2, 0D)]
    [InlineData("first-before", 1, 1, 20D, 1, 40D, 1, 20D)]
    [InlineData("last-after", 2, 1, 20D, 1, 40D, 2, 0D)]
    [InlineData("nested", 2, 2, 0D, 2, 20D, 1, 20D)]
    [InlineData("auto-wrapper", 2, 2, 0D, 2, 20D, 1, 20D)]
    public void HtmlConstrainedForcedBreak_PreservesAuthoredControlOriginAndFollowingFlow(
        string route, int pages, int secondPage, double secondTop, int thirdPage, double thirdTop,
        int afterPage, double afterTop) {
        string first = route is "after-first" or "adjacent" ? "break-after:page"
            : route == "first-before" ? "break-before:page" : "";
        string second = route is "before-second" or "nested" or "auto-wrapper" or "adjacent" ? "break-before:page"
            : route == "after-second" ? "break-after:page" : "";
        string third = route == "before-third" ? "break-before:page"
            : route == "last-after" ? "break-after:page" : "";
        string shortBox = "<div style='height:20px'><p style='" + first + "'>FORCED-ONE</p>"
            + "<p style='" + second + "'>FORCED-TWO</p><p style='" + third + "'>FORCED-THREE</p></div>";
        if (route == "auto-wrapper") shortBox = "<article>" + shortBox + "</article>";
        string content = shortBox + "<p style='padding-left:300px'>FORCED-AFTER</p>";
        if (route is "nested" or "auto-wrapper") content = "<section style='height:200px'>" + content + "</section>";

        foreach (HtmlRenderIntentProfile profile in new[] {
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(
                HtmlConversionDocument.Parse(FixedHeightOverflowSource(content)),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, FixedHeightOverflowOptions())).Document;

            Assert.Equal(pages, rendered.Pages.Count);
            AssertDescendantSpannerSourceOnce(rendered, "FORCED-ONE", "FORCED-TWO", "FORCED-THREE", "FORCED-AFTER");
            AssertFixedHeightOverflowPlacement(rendered, "FORCED-ONE", 1, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "FORCED-TWO", secondPage, secondTop);
            AssertFixedHeightOverflowPlacement(rendered, "FORCED-THREE", thirdPage, thirdTop);
            AssertFixedHeightOverflowPlacement(rendered, "FORCED-AFTER", afterPage, afterTop);
            rendered.RequireNoLoss();
        }
    }

    [Theory]
    [InlineData(1, 3)]
    [InlineData(2, 4)]
    public void HtmlConstrainedForcedBreak_EmptySiblingBoundariesRetainDistinctBlankPages(
        int emptySiblings, int pages) {
        string empty = string.Concat(Enumerable.Repeat(
            "<div style='break-before:page;break-after:page'></div>", emptySiblings));
        string html = FixedHeightOverflowSource("<div style='height:20px'><p>EMPTY-ONE</p>" + empty
            + "<p>EMPTY-TWO</p><p>EMPTY-THREE</p></div><p style='padding-left:300px'>EMPTY-AFTER</p>");

        foreach (HtmlRenderIntentProfile profile in new[] {
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, FixedHeightOverflowOptions())).Document;

            Assert.Equal(pages, rendered.Pages.Count);
            AssertDescendantSpannerSourceOnce(rendered, "EMPTY-ONE", "EMPTY-TWO", "EMPTY-THREE", "EMPTY-AFTER");
            AssertFixedHeightOverflowPlacement(rendered, "EMPTY-ONE", 1, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "EMPTY-AFTER", 1, 20D);
            AssertFixedHeightOverflowPlacement(rendered, "EMPTY-TWO", pages, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "EMPTY-THREE", pages, 20D);
            foreach (HtmlRenderPage blank in rendered.Pages.Skip(1).Take(emptySiblings))
                Assert.Empty(blank.Visuals.OfType<HtmlRenderText>());
            rendered.RequireNoLoss();
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlConstrainedForcedBreak_AfterControlRetainsBlankAllocationAndZeroHeightContent(
        bool blankAllocation) {
        string content = blankAllocation
            ? "<div style='height:20px'><div style='height:20px;break-after:page'></div>"
                + "<p>AFTER-ONE</p><p>AFTER-TWO</p></div>"
            : "<div style='height:0'><p style='break-after:page'>AFTER-ONE</p>"
                + "<p>AFTER-TWO</p><p>AFTER-THREE</p></div>";
        string html = FixedHeightOverflowSource(content + "<p style='padding-left:300px'>AFTER-FLOW</p>");

        foreach (HtmlRenderIntentProfile profile in new[] {
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, FixedHeightOverflowOptions())).Document;

            Assert.Equal(2, rendered.Pages.Count);
            AssertDescendantSpannerSourceOnce(rendered, blankAllocation
                ? new[] { "AFTER-ONE", "AFTER-TWO", "AFTER-FLOW" }
                : new[] { "AFTER-ONE", "AFTER-TWO", "AFTER-THREE", "AFTER-FLOW" });
            AssertFixedHeightOverflowPlacement(rendered, "AFTER-ONE", blankAllocation ? 2 : 1, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "AFTER-TWO", 2, blankAllocation ? 20D : 0D);
            AssertFixedHeightOverflowPlacement(rendered, "AFTER-FLOW", 2, 0D);
            if (blankAllocation) Assert.Empty(rendered.Pages[0].Visuals.OfType<HtmlRenderText>());
            else AssertFixedHeightOverflowPlacement(rendered, "AFTER-THREE", 2, 20D);
            rendered.RequireNoLoss();
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlConstrainedForcedBreak_AfterControlAdvancesFollowingFlowThroughConstrainedAncestors(
        bool terminal) {
        string html = FixedHeightOverflowSource("<section style='height:80px'><div style='height:20px'>"
            + "<p style='" + (terminal ? "" : "break-after:page") + "'>ANCESTOR-ONE</p>"
            + "<p>ANCESTOR-TWO</p><p style='" + (terminal ? "break-after:page" : "") + "'>ANCESTOR-THREE</p>"
            + "</div><p style='padding-left:300px'>ANCESTOR-INNER-AFTER</p></section>"
            + "<p style='padding-left:450px'>ANCESTOR-OUTER-AFTER</p>");

        foreach (HtmlRenderIntentProfile profile in new[] {
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, FixedHeightOverflowOptions())).Document;

            Assert.Equal(2, rendered.Pages.Count);
            AssertDescendantSpannerSourceOnce(rendered, "ANCESTOR-ONE", "ANCESTOR-TWO", "ANCESTOR-THREE",
                "ANCESTOR-INNER-AFTER", "ANCESTOR-OUTER-AFTER");
            AssertFixedHeightOverflowPlacement(rendered, "ANCESTOR-ONE", 1, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "ANCESTOR-TWO", terminal ? 1 : 2, terminal ? 20D : 0D);
            AssertFixedHeightOverflowPlacement(rendered, "ANCESTOR-THREE", terminal ? 1 : 2, terminal ? 40D : 20D);
            AssertFixedHeightOverflowPlacement(rendered, "ANCESTOR-INNER-AFTER", 2, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "ANCESTOR-OUTER-AFTER", 2, 0D);
            rendered.RequireNoLoss();
        }
    }
}
