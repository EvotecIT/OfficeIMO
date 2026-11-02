using AngleSharp.Html.Parser;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPageMarginContextCache_ManyUnusedNamesDoNotRetainEveryComputedCascade() {
        string css = string.Concat(Enumerable.Range(0, 40).Select(index =>
            "@page name" + index + "{@bottom-center{content:'x';font-family:Family" + index + "}}"));
        using var document = new HtmlParser().ParseDocument("<style>" + css + "</style><p>Body</p>");
        var options = PageMarginTypographyOptions();
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForRendering(document, options, limits);
        HtmlComputedStyle root = styles.Elements[document.DocumentElement!];
        HtmlCssPageRuleSet rules = HtmlCssPageSettingsResolver.Apply(document, options, new HtmlDiagnosticReport(), limits);

        var families = rules.EnumerateMarginFontRequests(root, styles, options, default)
            .Select(request => request.Families).ToHashSet(StringComparer.Ordinal);

        Assert.Contains("Family0", families);
        Assert.Contains("Family39", families);
        Assert.InRange(rules.CachedMarginContextCount, 1, 16);
    }

    [Fact]
    public void HtmlPageMarginContextCache_RepeatedGeneratedDeclarationsKeepLinearCandidateHistory() {
        var parent = new HtmlComputedStyle(new Dictionary<string, string>());
        var declarations = Enumerable.Range(0, 5_000).Select(index =>
            new HtmlGeneratedStyleDeclaration("font-size", index == 4_999 ? "20px" : "12px",
                false, 0, 0, null, index));

        HtmlComputedStyle result = HtmlComputedStyleEngine.ComputeGeneratedContext(
            parent, declarations, enforceResolutionLimits: true, default);

        Assert.Equal("20px", result.GetValue("font-size"));
    }

    [Fact]
    public void HtmlPageMarginContextCache_RepeatedPagesDoNotConsumeFreshRegisteredPropertyFanout() {
        string html = PageMarginContextCacheSource("white-space:pre;line-height:12px", "counter(page)")
            + "<p>" + string.Join("\n", Enumerable.Repeat("Body", 120)) + "</p>";
        var parsing = new HtmlConversionDocumentOptions { Trust = HtmlInputTrust.Untrusted };
        parsing.Limits.MaxCssDeclarations = 300;

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html, parsing),
            PageMarginTypographyOptions());

        Assert.True(rendered.Pages.Count >= 10);
        Assert.All(rendered.Pages, page => {
            HtmlRenderText margin = PageMarginTypographyText(page);
            Assert.Equal(page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture), margin.Text);
            Assert.Equal(12D, margin.Font.Size, 6);
        });
    }

    [Fact]
    public void HtmlPageMarginContextCache_FirstParityAndNamedMatchesKeepTheirOwnCascadeAndCounters() {
        const string html = """
            <style>
              html { font:12px/1 Pinned; }
              body,p { margin:0; font:9px/12px Pinned; }
              p+p { break-before:page; }
              @page { size:300px 240px; margin:60px; @bottom-center { content:counter(page); } }
              @page :first { color:red; }
              @page :left { color:green; }
              @page :right { color:blue; }
              @page invoice { font-size:20px; }
              @page invoice:left { color:orange; }
            </style><p>First</p><p>Second</p><p>Third</p>
            <p style='page:invoice'>Fourth</p><p style='page:invoice'>Fifth</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html),
            PageMarginTypographyOptions());

        Assert.Equal(5, rendered.Pages.Count);
        OfficeColor[] colors = { OfficeColor.Red, OfficeColor.Green, OfficeColor.Blue,
            OfficeColor.FromRgb(255, 165, 0), OfficeColor.Blue };
        for (int index = 0; index < rendered.Pages.Count; index++) {
            HtmlRenderText margin = PageMarginTypographyText(rendered.Pages[index]);
            Assert.Equal((index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), margin.Text);
            Assert.Equal(colors[index], margin.Color);
            Assert.Equal(index < 3 ? 12D : 20D, margin.Font.Size, 6);
        }
    }

    [Fact]
    public void HtmlPageMarginContextCache_RootAndComputedStyleScopesKeepTheirValuesAndBudgets() {
        using var document = new HtmlParser().ParseDocument(PageMarginContextCacheSource(string.Empty, "'A'") + "<p>Body</p>");
        var options = PageMarginTypographyOptions();
        var limits = HtmlConversionLimits.CreateTrustedProfile();
        limits.MaxCssDeclarations = 300;
        HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForRendering(document, options, limits);
        HtmlComputedStyle root = styles.Elements[document.DocumentElement!];
        HtmlCssPageRuleSet rules = HtmlCssPageSettingsResolver.Apply(document, options, new HtmlDiagnosticReport(), limits);

        Assert.All(rules.EnumerateMarginFontRequests(root, styles, options, default),
            request => Assert.Equal("Pinned", request.Families));
        var changedProperties = root.Properties.ToDictionary(pair => pair.Key, pair => pair.Value);
        changedProperties["font-family"] = "Other";
        var changedRoot = new HtmlComputedStyle(changedProperties);
        Assert.All(rules.EnumerateMarginFontRequests(changedRoot, styles, options, default),
            request => Assert.Equal("Other", request.Families));

        // A fresh style operation has its own registration budget even when it
        // reaches the same parsed page rules and root snapshot.
        limits.MaxCssDeclarations = 100;
        HtmlComputedStyleSet restricted = HtmlComputedStyleEngine.ComputeForRendering(document, options, limits);
        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() =>
            rules.EnumerateMarginFontRequests(root, restricted, options, default).ToArray());
        Assert.Equal("MaxCssDeclarations", error.LimitSource);
    }

    [Fact]
    public void HtmlPageMarginContextCache_HitsRecomputeUsedGeometryAndHonorCancellation() {
        using var document = new HtmlParser().ParseDocument("<style>html{font:12px/1 Pinned}</style><p>Body</p>");
        var options = PageMarginTypographyOptions();
        HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForRendering(document, options,
            HtmlConversionLimits.CreateTrustedProfile());
        var rules = new HtmlCssPageRuleSet();
        rules.Add(new HtmlCssPageRule(null, HtmlCssPageSelector.Generic,
            new Dictionary<HtmlCssPageMarginPosition, HtmlCssPageMarginRule> {
                [HtmlCssPageMarginPosition.BottomCenter] = new HtmlCssPageMarginRule(
                    new HtmlCssPageDeclaration("'A'", false, 0), Array.Empty<HtmlGeneratedStyleDeclaration>())
            }, default, default, null, new[] {
                new HtmlGeneratedStyleDeclaration("font-size", "10vw", false, 0, 0, null, 0)
            }));
        using var cancellation = new CancellationTokenSource();
        var resolver = new HtmlRenderStyleResolver(styles, options, new HtmlDiagnosticReport(), _ => double.NaN,
            cancellation.Token);

        HtmlCssPageMarginTemplate narrow = Assert.Single(rules.ResolveMarginBoxes(2, null,
            new HtmlCssPageGeometry(240D, 360D, HtmlRenderMargins.All(60D)), options, resolver)).Value;
        HtmlCssPageMarginTemplate wide = Assert.Single(rules.ResolveMarginBoxes(4, null,
            new HtmlCssPageGeometry(480D, 360D, HtmlRenderMargins.All(60D)), options, resolver)).Value;

        Assert.Equal(24D, narrow.Font.Size, 6);
        Assert.Equal(48D, wide.Font.Size, 6);
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => rules.ResolveMarginBoxes(6, null,
            new HtmlCssPageGeometry(480D, 360D, HtmlRenderMargins.All(60D)), options, resolver));
    }

    private static string PageMarginContextCacheSource(string bodyDeclarations, string content) =>
        "<style>" + string.Concat(Enumerable.Range(0, 20).Select(index =>
            "@property --n" + index + "{syntax:'<length>';inherits:false;initial-value:12px}"))
        + "html{font:12px/1 Pinned}body,p{margin:0;" + bodyDeclarations + "}"
        + "@page{size:300px 240px;margin:60px;@bottom-center{content:" + content
        + ";font-size:var(--n0)}}</style>";
}
