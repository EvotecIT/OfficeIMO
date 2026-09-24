using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPseudoRulesDoNotSpendSelectorBudgetOnOrdinaryRules() {
        string ordinaryRules = string.Concat(Enumerable.Range(0, 10)
            .Select(index => $":not(.missing-{index}){{color:#123456}}"));
        string html = "<style>" + ordinaryRules + "p::before{content:'BudgetMarker'}</style><p>Body</p>";
        var options = HtmlConversionDocumentOptions.CreateUntrustedProfile();
        HtmlConversionLimits limits = options.Limits.Clone();
        limits.MaxSelectorEvaluations = 300;
        options.Limits = limits;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html, options));

        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(),
            text => text.Text == "BudgetMarker" && text.Source == "p::before");
    }

    [Fact]
    public void HtmlComplexSelectorsReuseStableAncestorMatchesWithinBudget() {
        string children = string.Concat(Enumerable.Repeat("<p>Item</p>", 100));
        var options = HtmlConversionDocumentOptions.CreateUntrustedProfile();
        HtmlConversionLimits limits = options.Limits.Clone();
        limits.MaxSelectorEvaluations = 150;
        options.Limits = limits;
        options.IncludeNormalizedHtml = false;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>.wrapper p{color:rgb(1, 2, 3)}</style><div class='wrapper'>" + children + "</div>", options);

        Assert.Contains("rgba(1, 2, 3, 1)", document.StyleSummary.ColorValues);
    }
}
