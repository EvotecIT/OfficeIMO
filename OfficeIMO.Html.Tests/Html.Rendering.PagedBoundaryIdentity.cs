using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("block", false)]
    [InlineData("block", true)]
    [InlineData("column-flex", false)]
    [InlineData("column-flex", true)]
    [InlineData("column-flex-gap", false)]
    [InlineData("column-flex-gap", true)]
    [InlineData("inline", false)]
    [InlineData("inline", true)]
    public void HtmlPagedBoundary_NamedPageAndAuthoredForceShareOneSiblingTransition(string host, bool followingBefore) {
        const string one = "<p style='break-after:page'>ONE</p>";
        string two = "<p style='page:other" + (followingBefore ? ";break-before:page" : "") + "'>TWO</p>";
        string body = host switch {
            "column-flex" => "<div style='display:flex;flex-direction:column'>" + one + two + "</div>",
            "column-flex-gap" => "<div style='display:flex;flex-direction:column;gap:12px'>" + one + two + "</div>",
            "inline" => "<span>" + one + two + "</span>",
            _ => "<div>" + one + two + "</div>"
        };
        string html = ColumnBreakSource("<style>@page{size:700px 140px;margin:20px}"
            + "@page other{size:700px 140px;margin:20px}</style>" + body);

        foreach (HtmlRenderIntentProfile profile in new[] { HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, ColumnPageBandOptions())).Document;

            Assert.Equal(2, rendered.Pages.Count);
            Assert.Null(rendered.Pages[0].PageName);
            Assert.Equal("other", rendered.Pages[1].PageName);
            Assert.Equal(1, ColumnPageBandTextPage(rendered, "ONE").PageNumber);
            Assert.Equal(2, ColumnPageBandTextPage(rendered, "TWO").PageNumber);
            AssertColumnPageBandTextOnce(rendered, new[] { "ONE", "TWO" });
            rendered.RequireNoLoss();
        }
    }

    [Theory]
    [InlineData("block", "right", null, 3)]
    [InlineData("block", "right", "page", 3)]
    [InlineData("block", "right", "left", 2)]
    [InlineData("block", "left", "right", 3)]
    [InlineData("block", "left", "page", 2)]
    [InlineData("column-flex-gap", "right", null, 3)]
    [InlineData("column-flex-gap", "right", "page", 3)]
    [InlineData("inline", "right", null, 3)]
    public void HtmlPagedBoundary_NamedTransitionRetainsTheLatestAuthoredPageSide(
        string host, string after, string? before, int expectedPages) {
        string children = "<p style='break-after:" + after + "'>ONE</p><p style='page:other"
            + (before == null ? "" : ";break-before:" + before) + "'>TWO</p>";
        string body = host switch {
            "column-flex-gap" => "<div style='display:flex;flex-direction:column;gap:12px'>" + children + "</div>",
            "inline" => "<span>" + children + "</span>",
            _ => "<div>" + children + "</div>"
        };
        string html = ColumnBreakSource("<style>@page{size:700px 140px;margin:20px}"
            + "@page other{size:700px 140px;margin:20px}</style>" + body);

        foreach (HtmlRenderIntentProfile profile in new[] { HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, ColumnPageBandOptions())).Document;

            Assert.Equal(expectedPages, rendered.Pages.Count);
            Assert.Equal(1, ColumnPageBandTextPage(rendered, "ONE").PageNumber);
            Assert.Equal(expectedPages, ColumnPageBandTextPage(rendered, "TWO").PageNumber);
            Assert.Equal("other", rendered.Pages[expectedPages - 1].PageName);
            if (expectedPages == 3) Assert.Empty(rendered.Pages[1].Visuals.OfType<HtmlRenderText>());
            AssertColumnPageBandTextOnce(rendered, new[] { "ONE", "TWO" });
            rendered.RequireNoLoss();
        }
    }

    [Theory]
    [InlineData("right", true)]
    [InlineData("left", false)]
    public void HtmlPagedBoundary_NamedPageSideCountsBlankPagesAgainstTheLimit(string side, bool exceedsLimit) {
        string html = ColumnBreakSource("<style>@page{size:700px 140px;margin:20px}"
            + "@page other{size:700px 140px;margin:20px}</style>"
            + "<div><p style='break-after:" + side + "'>ONE</p><p style='page:other'>TWO</p></div>");
        foreach (HtmlRenderIntentProfile profile in new[] { HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
            HtmlRenderOptions options = ColumnPageBandOptions();
            options.MaxPageCount = 2;
            HtmlRenderDocument Render() => HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, options)).Document;
            if (exceedsLimit) {
                InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => Render());
                Assert.Contains("maximum page count", error.Message, StringComparison.OrdinalIgnoreCase);
            } else {
                HtmlRenderDocument rendered = Render();
                Assert.Equal(2, rendered.Pages.Count);
                Assert.Equal(2, ColumnPageBandTextPage(rendered, "TWO").PageNumber);
                rendered.RequireNoLoss();
            }
        }
    }
}
