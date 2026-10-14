using System.Text;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlAllSeverityFinalHighSecurityTests {
    [Theory]
    [InlineData("empty")]
    [InlineData("cached-text")]
    [InlineData("nested")]
    public void TableIntrinsicInlineDescendantsConsumeTheLayoutOperationBudget(string content) {
        string descendants = content switch {
            "cached-text" => string.Concat(Enumerable.Repeat("<span>Cached</span>", 64)),
            "nested" => string.Concat(Enumerable.Repeat("<span><span></span></span>", 32)),
            _ => string.Concat(Enumerable.Repeat("<span></span>", 64))
        };
        string html = "<table><tr><td>" + descendants + "</td></tr></table>";
        Assert.Single(HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { MaxLayoutOperations = 4096 }).Pages);

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { MaxLayoutOperations = 8 }));

        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), exception.LimitSource);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void TableDescendantIntrinsicSizingConsumesTheLayoutOperationBudget() {
        var html = new StringBuilder("<table><tr><td>");
        for (int depth = 0; depth < 8; depth++) {
            html.Append("<div>");
        }
        for (int image = 0; image < 12; image++) {
            html.Append("<img width='1' height='1'>");
        }
        for (int depth = 0; depth < 8; depth++) {
            html.Append("</div>");
        }
        html.Append("</td></tr></table>");

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render(
                html.ToString(),
                new HtmlRenderOptions {
                    MaxLayoutDepth = 32,
                    MaxLayoutOperations = 24
                }));

        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), exception.LimitSource);
    }
}
