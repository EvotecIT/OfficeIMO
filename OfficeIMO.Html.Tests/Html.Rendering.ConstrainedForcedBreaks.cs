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
}
