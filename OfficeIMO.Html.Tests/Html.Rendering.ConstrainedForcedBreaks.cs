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

    [Theory]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(1, true)]
    public void HtmlConstrainedForcedBreak_InitialEmptyRunMovesOnceWithoutOwnerEntryRecursion(
        int emptySiblings, bool followingBefore) {
        string empty = string.Concat(Enumerable.Repeat(
            "<div style='break-before:page;break-after:page'></div>", emptySiblings));
        string html = FixedHeightOverflowSource("<div style='height:20px'>" + empty
            + "<p style='" + (followingBefore ? "break-before:page" : "") + "'>INITIAL-ONE</p>"
            + "<p>INITIAL-TWO</p></div><p style='padding-left:300px'>INITIAL-AFTER</p>");

        foreach (HtmlRenderIntentProfile profile in new[] {
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, FixedHeightOverflowOptions())).Document;

            Assert.Equal(2, rendered.Pages.Count);
            Assert.Empty(rendered.Pages[0].Visuals.OfType<HtmlRenderText>());
            AssertDescendantSpannerSourceOnce(rendered, "INITIAL-ONE", "INITIAL-TWO", "INITIAL-AFTER");
            AssertFixedHeightOverflowPlacement(rendered, "INITIAL-ONE", 2, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "INITIAL-TWO", 2, 20D);
            AssertFixedHeightOverflowPlacement(rendered, "INITIAL-AFTER", 2, 0D);
            rendered.RequireNoLoss();
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlConstrainedForcedBreak_ZeroHeightPrecedingPaintSurvivesItsAuthoredAfterBoundary(
        bool followingBefore) {
        string html = FixedHeightOverflowSource("<div style='height:20px'>"
            + "<p style='height:0;break-after:page'>VISIBLE-ONE</p>"
            + "<p style='" + (followingBefore ? "break-before:page" : "") + "'>VISIBLE-TWO</p>"
            + "<p>VISIBLE-THREE</p></div><p style='padding-left:300px'>VISIBLE-AFTER</p>");

        foreach (HtmlRenderIntentProfile profile in new[] {
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, FixedHeightOverflowOptions())).Document;

            Assert.Equal(2, rendered.Pages.Count);
            AssertDescendantSpannerSourceOnce(rendered, "VISIBLE-ONE", "VISIBLE-TWO", "VISIBLE-THREE", "VISIBLE-AFTER");
            AssertFixedHeightOverflowPlacement(rendered, "VISIBLE-ONE", 1, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "VISIBLE-TWO", 2, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "VISIBLE-THREE", 2, 20D);
            AssertFixedHeightOverflowPlacement(rendered, "VISIBLE-AFTER", 2, 0D);
            rendered.RequireNoLoss();
        }
    }

    [Fact]
    public void HtmlConstrainedForcedBreak_InteriorSideAtZeroRetainsDiagnosedFallback() {
        string html = FixedHeightOverflowSource("<div style='height:20px'>"
            + "<div style='break-after:left'></div><p>SIDE-ONE</p><p>SIDE-TWO</p></div>"
            + "<p style='padding-left:300px'>SIDE-AFTER</p>");

        foreach (HtmlRenderIntentProfile profile in new[] {
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, FixedHeightOverflowOptions())).Document;

            Assert.Contains(rendered.Diagnostics,
                diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ConstrainedOverflowPaginationPending);
            Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlConstrainedForcedBreak_AnonymousFollowingRunReplaysByOriginalNodeIdentity(bool inlineElement) {
        string following = inlineElement ? "<span>ANONYMOUS-TWO</span>" : "ANONYMOUS-TWO";
        string html = FixedHeightOverflowSource("<div style='height:20px'>"
            + "<p style='height:0;break-after:page'>ANONYMOUS-ONE</p>" + following
            + "<p>ANONYMOUS-THREE</p></div><p style='padding-left:300px'>ANONYMOUS-AFTER</p>");

        foreach (HtmlRenderIntentProfile profile in new[] {
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, FixedHeightOverflowOptions())).Document;

            Assert.Equal(2, rendered.Pages.Count);
            AssertDescendantSpannerSourceOnce(rendered,
                "ANONYMOUS-ONE", "ANONYMOUS-TWO", "ANONYMOUS-THREE", "ANONYMOUS-AFTER");
            AssertFixedHeightOverflowPlacement(rendered, "ANONYMOUS-ONE", 1, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "ANONYMOUS-TWO", 2, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "ANONYMOUS-THREE", 2, 20D);
            AssertFixedHeightOverflowPlacement(rendered, "ANONYMOUS-AFTER", 2, 0D);
            rendered.RequireNoLoss();
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(8)]
    public void HtmlConstrainedForcedBreak_TerminalAfterRetainsOwnedPaintAndIndependentFollowingAllocation(int lines) {
        string[] labels = Enumerable.Range(1, lines).Select(index => "TERMINAL-" + index).ToArray();
        string html = FixedHeightOverflowSource("<div style='height:20px'>"
            + "<p style='height:0;break-after:page'>" + string.Join("<br>", labels)
            + "</p></div><p style='padding-left:300px'>TERMINAL-AFTER</p>");

        foreach (HtmlRenderIntentProfile profile in new[] {
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, FixedHeightOverflowOptions())).Document;

            Assert.Equal(2, rendered.Pages.Count);
            AssertDescendantSpannerSourceOnce(rendered, labels.Concat(new[] { "TERMINAL-AFTER" }).ToArray());
            for (int index = 0; index < labels.Length; index++)
                AssertFixedHeightOverflowPlacement(rendered, labels[index], index / 5 + 1, index % 5 * 20D);
            // The after boundary belongs to the first destination edge even
            // when independent visible overflow continues farther down page two.
            AssertFixedHeightOverflowPlacement(rendered, "TERMINAL-AFTER", 2, 0D);
            rendered.RequireNoLoss();
        }
    }

    [Theory]
    [InlineData(HtmlRenderUserAgentStyleMode.Document, false)]
    [InlineData(HtmlRenderUserAgentStyleMode.Document, true)]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser, false)]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser, true)]
    public void HtmlConstrainedForcedBreak_FittingEntryControlKeepsOneSharedFollowingFlow(
        HtmlRenderUserAgentStyleMode userAgent, bool transparentRoot) {
        string html = FixedHeightOverflowSource("<body" + (transparentRoot ? " style='display:contents'" : "") + ">"
            + "<div style='height:20px'>PREFIX</div><div style='height:20px'>"
            + "<div style='break-after:page'></div><p>ENTRY-ONE</p></div>"
            + "<p style='padding-left:300px'>ENTRY-AFTER</p></body>").Replace(";background:white", "");

        foreach (HtmlRenderIntentProfile profile in new[] {
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderOptions options = FixedHeightOverflowOptions();
            options.UserAgentStyles = userAgent;
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, options)).Document;

            Assert.Equal(2, rendered.Pages.Count);
            AssertDescendantSpannerSourceOnce(rendered, "PREFIX", "ENTRY-ONE", "ENTRY-AFTER");
            AssertFixedHeightOverflowPlacement(rendered, "PREFIX", 1, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "ENTRY-ONE", 2, 0D);
            AssertFixedHeightOverflowPlacement(rendered, "ENTRY-AFTER", 2, 0D);
            rendered.RequireNoLoss();
        }
    }
}
