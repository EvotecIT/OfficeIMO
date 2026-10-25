using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPageMarginFontPreparation_ReferencedRunningValuesShareTheLayoutBudget(bool attributeValue) {
        string html = PageMarginFontBudgetSource(referenced: true, attributeValue: attributeValue);
        AssertPageMarginFontBudget(html);
    }

    [Theory]
    [InlineData("absolute")]
    [InlineData("fixed")]
    public void HtmlPageMarginFontPreparation_GlobalPositionedAssignmentsShareTheLayoutBudget(string position) {
        string html = PageMarginFontBudgetSource(referenced: true, attributeValue: false)
            .Replace("</style>", "#assignment{position:" + position + ";left:0;top:0}</style>");

        AssertPageMarginFontBudget(html);
    }

    [Fact]
    public void HtmlPageMarginFontPreparation_UnreferencedAssignmentsDoNotConsumeCoverageWork() {
        string html = PageMarginFontBudgetSource(referenced: false, attributeValue: false);

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html),
            new HtmlToPdfOptions { MaxLayoutOperations = 3_500, AllowSystemFontFallback = false });

        Assert.Single(rendered.Pages);
        Assert.Equal("BODY", rendered.Text.Trim());
    }

    private static void AssertPageMarginFontBudget(string html) {
        var options = new HtmlToPdfOptions { MaxLayoutOperations = 3_500, AllowSystemFontFallback = false };
        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options));

        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, error.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), error.LimitSource);

        options.MaxLayoutOperations = 100_000;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
        Assert.Single(rendered.Pages);
        Assert.Equal("BODY", rendered.Text.Trim());
    }

    private static string PageMarginFontBudgetSource(bool referenced, bool attributeValue) {
        // ASCII needs no installed font. Distinct used font sizes require four
        // coverage inputs while the assignment is parsed only once.
        string value = new string('A', 1_024);
        string content = referenced ? "string(heading)" : "'ASCII'";
        string declaration = attributeValue ? "attr(data-heading)" : "'" + value + "'";
        return "<style>@page{size:300px 200px;margin:40px;"
            + "@top-left{content:" + content + ";font:8px Arial}"
            + "@top-right{content:" + content + ";font:9px Arial}"
            + "@bottom-left{content:" + content + ";font:10px Arial}"
            + "@bottom-right{content:" + content + ";font:11px Arial}}"
            + "body{margin:0}#assignment{string-set:heading " + declaration + "}</style>"
            + "<main id='assignment'" + (attributeValue ? " data-heading='" + value + "'" : string.Empty) + ">BODY</main>";
    }
}
