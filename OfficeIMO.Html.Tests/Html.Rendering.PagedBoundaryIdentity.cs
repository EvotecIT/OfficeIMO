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
}
