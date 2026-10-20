using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("<div style='break-after:page'><p style='break-after:right'>ONE</p></div><p>TWO</p>", 3)]
    [InlineData("<div style='break-after:right'><p style='break-after:page'>ONE</p></div><p>TWO</p>", 3)]
    [InlineData("<p>ONE</p><div style='break-before:right'><p style='break-before:page'>TWO</p></div>", 3)]
    [InlineData("<p style='break-after:right'>ONE</p><p style='page:other;break-before:left'>TWO</p>", 2)]
    [InlineData("<p style='break-after:left'>ONE</p><p style='page:other;break-before:right'>TWO</p>", 3)]
    [InlineData("<p style='break-after:right'>ONE</p><p style='page:other;break-before:page'>TWO</p>", 3)]
    public void HtmlPagedBoundary_RootSiblingsCombineParentAndDescendantPageControls(string body, int expectedPages) {
        string html = ColumnBreakSource("<style>@page{size:700px 140px;margin:20px}"
            + "@page other{size:700px 140px;margin:20px}</style>" + body);
        foreach (HtmlRenderUserAgentStyleMode userAgent in new[] { HtmlRenderUserAgentStyleMode.Document, HtmlRenderUserAgentStyleMode.Browser }) {
            foreach (HtmlRenderIntentProfile profile in new[] { HtmlRenderIntentProfile.PrintPaged, HtmlRenderIntentProfile.ScreenMediaPaged }) {
                HtmlRenderOptions options = ColumnPageBandOptions();
                options.UserAgentStyles = userAgent;
                options.MaxPageCount = expectedPages;
                HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                    HtmlRenderRequest.Create(profile, HtmlRenderEncoder.DisplayList, options)).Document;

                Assert.Equal(expectedPages, rendered.Pages.Count);
                Assert.Equal(1, ColumnPageBandTextPage(rendered, "ONE").PageNumber);
                Assert.Equal(expectedPages, ColumnPageBandTextPage(rendered, "TWO").PageNumber);
                if (body.Contains("page:other", StringComparison.Ordinal))
                    Assert.Equal("other", rendered.Pages[expectedPages - 1].PageName);
                AssertColumnPageBandTextOnce(rendered, new[] { "ONE", "TWO" });
                rendered.RequireNoLoss();
            }
        }
    }
}
