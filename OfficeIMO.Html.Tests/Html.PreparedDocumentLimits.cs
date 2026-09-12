using System.Linq;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPreparedDocumentLimitTests {
    [Fact]
    public void PreparedStylesPreserveUnboundedContractWhileRawInputRemainsBounded() {
        int rules = HtmlConversionLimits.CreateUntrustedProfile().MaxCssRules!.Value + 1;
        string source = "<style>" + string.Concat(Enumerable.Repeat(".unused{color:red}", rules)) + "p{color:blue}</style><p>Text</p>";
        HtmlDocument owned = AngleSharpHtmlParser.Instance.Parse(source, new HtmlParseOptions());
        HtmlElement paragraph = owned.QuerySelector("p")!;
        var computed = HtmlComputedStyleEngine.Compute(owned);
        Assert.Contains("0, 0, 255", computed[paragraph].GetValue("color"));
        Assert.Equal("MaxCssRules", Assert.Throws<HtmlDomLimitException>(() => HtmlComputedStyleEngine.Compute(source)).LimitSource);
    }

    [Fact]
    public void PreparedConversionStylesUseTheCallersCssAndTreeBudgets() {
        const string source = "<style>p{color:blue} p{font-weight:bold}</style><p>Text</p>";
        var limited = HtmlConversionDocument.Parse(source, new HtmlConversionDocumentOptions {
            Limits = new HtmlConversionLimits { MaxCssRules = 1 }
        });
        Assert.Equal("MaxCssRules", Assert.Throws<HtmlDomLimitException>(() => HtmlComputedStyleEngine.Compute(limited)).LimitSource);
        var trusted = HtmlConversionDocument.FromDocument(limited.Document, HtmlConversionDocumentOptions.CreateTrustedProfile());
        Assert.Equal("bold", HtmlComputedStyleEngine.Compute(trusted)[trusted.Document.QuerySelector("p")!].GetValue("font-weight"));
        Assert.Equal("MaxHtmlNodes", Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.FromDocument(limited.Document,
            new HtmlConversionDocumentOptions { Limits = new HtmlConversionLimits { MaxHtmlNodes = 2 } })).LimitSource);
    }
}
