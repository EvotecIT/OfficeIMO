using System.Text;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlAllSeverityFinalHighSecurityTests {
    [Fact]
    public void TableImageAncestorStyleWalksConsumeTheLayoutOperationBudget() {
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
                    MaxLayoutOperations = 80
                }));

        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), exception.LimitSource);
    }
}
