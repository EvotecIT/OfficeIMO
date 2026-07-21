using OfficeIMO.Html;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlConversionLimitTests {
    [Fact]
    public void HtmlConversionOptions_TrustSelectsOneConsistentDefaultProfile() {
        var untrusted = new HtmlConversionDocumentOptions();
        var trusted = new HtmlConversionDocumentOptions { Trust = HtmlInputTrust.Trusted };

        Assert.True(untrusted.UrlPolicy.DisallowFileUrls);
        Assert.False(untrusted.UrlPolicy.AllowDataUrls);
        Assert.NotNull(untrusted.Limits.MaxHtmlNodes);
        Assert.Equal(64, untrusted.Limits.MaxResponsiveImageCandidates);

        Assert.False(trusted.UrlPolicy.DisallowFileUrls);
        Assert.True(trusted.UrlPolicy.AllowDataUrls);
        Assert.Null(trusted.Limits.MaxHtmlNodes);
        Assert.Null(trusted.Limits.MaxResponsiveImageCandidates);
    }

    [Fact]
    public void HtmlConversionOptions_ExplicitPolicyAndLimitsSurviveTrustChanges() {
        var explicitPolicy = HtmlUrlPolicy.CreateHyperlinkProfile();
        var explicitLimits = new HtmlConversionLimits { MaxHtmlNodes = 17 };
        var options = new HtmlConversionDocumentOptions {
            UrlPolicy = explicitPolicy,
            Limits = explicitLimits,
            Trust = HtmlInputTrust.Trusted
        };

        Assert.Same(explicitPolicy, options.UrlPolicy);
        Assert.Same(explicitLimits, options.Limits);
        Assert.Equal(17, options.Limits.MaxHtmlNodes);
        Assert.True(options.UrlPolicy.DisallowFileUrls);
    }

    [Fact]
    public void HtmlConversionOptions_ResourcePolicyDoesNotDependOnInitializerOrder() {
        var trustFirst = new HtmlConversionDocumentOptions {
            Trust = HtmlInputTrust.Trusted,
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        };
        var policyFirst = new HtmlConversionDocumentOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            Trust = HtmlInputTrust.Trusted
        };

        Assert.True(trustFirst.ResourceUrlPolicy.RestrictUrlSchemes);
        Assert.True(policyFirst.ResourceUrlPolicy.RestrictUrlSchemes);
        Assert.True(trustFirst.ResourceUrlPolicy.DisallowFileUrls);
        Assert.True(policyFirst.ResourceUrlPolicy.DisallowFileUrls);
        Assert.Equal(trustFirst.ResourceUrlPolicy.AllowedUrlSchemes.OrderBy(value => value),
            policyFirst.ResourceUrlPolicy.AllowedUrlSchemes.OrderBy(value => value));
    }

    [Fact]
    public void HtmlConversionOptions_RejectUnknownTrustAndProfileValues() {
        Assert.Throws<ArgumentOutOfRangeException>(() => HtmlConversionDocument.Parse(
            "<p>x</p>",
            new HtmlConversionDocumentOptions { Trust = (HtmlInputTrust)999 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => HtmlConversionDocument.Parse(
            "<p>x</p>",
            new HtmlConversionDocumentOptions { Profile = (HtmlConversionProfile)999 }));
    }

    [Fact]
    public void HtmlConversionDocument_RejectsSourceAndDomBeforeAnalysis() {
        HtmlDomLimitException sourceException = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlConversionDocument.Parse("<p>too long</p>", Options(limits => limits.MaxInputCharacters = 8)));
        HtmlDomLimitException nodeException = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlConversionDocument.Parse("<main><p>one</p><p>two</p></main>", Options(limits => limits.MaxHtmlNodes = 4)));
        HtmlDomLimitException depthException = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlConversionDocument.Parse("<main><section><p>deep</p></section></main>", Options(limits => limits.MaxHtmlDepth = 3)));

        Assert.Equal(HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded, sourceException.Code);
        Assert.Equal(HtmlRenderDiagnosticCodes.NodeLimitExceeded, nodeException.Code);
        Assert.Equal(HtmlConversionDiagnosticCodes.HtmlDepthLimitExceeded, depthException.Code);
    }

    [Fact]
    public void HtmlConversionDocument_EnforcesEmbeddedCssBytesBeforeAnalysis() {
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxCssBytes = 8;
        limits.MaxTotalCssBytes = 16;

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlConversionDocument.Parse("<style>.a{color:red}</style><p class='a'>x</p>", new HtmlConversionDocumentOptions { Limits = limits }));

        Assert.Equal(HtmlConversionDiagnosticCodes.CssSizeLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlConversionLimits.MaxCssBytes), exception.LimitSource);
    }

    [Fact]
    public void HtmlConversionDocument_EnforcesSharedDomAndCssLimitsInsideSrcDoc() {
        const string html = "<iframe srcdoc=\"<style>.a{color:red}</style><main><p>x</p></main>\"></iframe>";
        HtmlDomLimitException nodeException = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlConversionDocument.Parse(html, Options(limits => limits.MaxHtmlNodes = 6)));

        var cssLimits = HtmlConversionLimits.CreateUntrustedProfile();
        cssLimits.MaxCssBytes = 8;
        cssLimits.MaxTotalCssBytes = 16;
        HtmlDomLimitException cssException = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions { Limits = cssLimits }));

        Assert.Equal(HtmlRenderDiagnosticCodes.NodeLimitExceeded, nodeException.Code);
        Assert.Equal(HtmlConversionDiagnosticCodes.CssSizeLimitExceeded, cssException.Code);
    }

    [Fact]
    public void RawNormalizerAndResourceOverloadsHonorCallerSourceLimitsBeforeParsing() {
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxInputCharacters = 8;

        HtmlDomLimitException normalizeException = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlNormalizer.Normalize("<p>too long</p>", new HtmlNormalizationOptions { Limits = limits }));
        HtmlDomLimitException resourceException = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlResourcePipeline.BuildManifest("<p>too long</p>", new HtmlResourcePipelineOptions { Limits = limits }));

        Assert.Equal(HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded, normalizeException.Code);
        Assert.Equal(HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded, resourceException.Code);
    }

    [Fact]
    public void HtmlConversionDocument_EnforcesSemanticAttributeSizeBeforeAdapters() {
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxSemanticMetadataCharacters = 4;

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlConversionDocument.Parse(
                "<main data-officeimo-profile='oversized'></main>",
                new HtmlConversionDocumentOptions { Limits = limits }));

        Assert.Equal(HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlConversionLimits.MaxSemanticMetadataCharacters), exception.LimitSource);
    }

    [Fact]
    public void HtmlConversionDocument_UsesOneResponsiveCandidateLimitForNormalizationAndResources() {
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxResponsiveImageCandidates = 1;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<img srcset='https://example.test/first.png 1x, https://example.test/second.png 2x'>",
            new HtmlConversionDocumentOptions { Limits = limits });

        Assert.Contains("first.png", document.HtmlForConversion, StringComparison.Ordinal);
        Assert.DoesNotContain("second.png", document.HtmlForConversion, StringComparison.Ordinal);
        Assert.Single(document.ResourceManifest.Resources, resource => resource.Source.Contains("first.png", StringComparison.Ordinal));
        Assert.DoesNotContain(document.ResourceManifest.Resources, resource => resource.Source.Contains("second.png", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlConversionDocument_UsesSeparateHyperlinkAndResourcePoliciesInTheManifest() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<a href='data:text/plain,blocked'>link</a><img src='data:image/png;base64,AQID'>");

        HtmlResourceReference hyperlink = Assert.Single(document.ResourceManifest.Resources,
            resource => resource.Kind == HtmlResourceKind.Hyperlink);
        HtmlResourceReference image = Assert.Single(document.ResourceManifest.Resources,
            resource => resource.Kind == HtmlResourceKind.Image);

        Assert.False(hyperlink.IsAllowed);
        Assert.True(image.IsAllowed);
    }

    [Fact]
    public void HtmlConversionDocument_DefersStyleWorkAndEnforcesItsComplexityBudgetOnDemand() {
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxCssRules = 1;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>.a{color:red}.b{color:blue}</style><p class='a'>x</p>",
            new HtmlConversionDocumentOptions { Limits = limits, IncludeNormalizedHtml = false });

        Assert.Contains("class='a'", document.SourceHtml);
        Assert.NotNull(document.LogicalDocument);
        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() => _ = document.StyleSummary);
        Assert.Equal(HtmlConversionDiagnosticCodes.CssRuleLimitExceeded, exception.Code);
    }

    [Fact]
    public void HtmlConversionDocument_LoadStopsAtTheSharedCharacterBudgetAndRestoresTheStream() {
        byte[] bytes = Encoding.UTF8.GetBytes("<p>stream content</p>");
        using var stream = new MemoryStream(bytes);
        stream.Position = 3;

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlConversionDocument.Load(stream, Options(limits => limits.MaxInputCharacters = 8)));

        Assert.Equal(HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded, exception.Code);
        Assert.Equal(3, stream.Position);
    }

    [Fact]
    public void HtmlComputedStyles_IndexIrrelevantSelectorsWithoutChangingTheCascade() {
        string irrelevantRules = string.Concat(Enumerable.Range(0, 200).Select(index => ".unused" + index + "{color:red}"));
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxSelectorEvaluations = 8;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>" + irrelevantRules + "#target{color:rgb(1, 2, 3)}</style><p id='target'>indexed</p>",
            new HtmlConversionDocumentOptions { Limits = limits, IncludeNormalizedHtml = false });

        Assert.Contains("rgba(1, 2, 3, 1)", document.StyleSummary.ColorValues);
    }

    [Fact]
    public void HtmlComputedStyles_KeepEscapedIdentifiersOnTheConservativeSelectorPath() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>#escaped\\:id{color:rgb(4, 5, 6)}</style><p id='escaped:id'>escaped</p>",
            new HtmlConversionDocumentOptions { IncludeNormalizedHtml = false });

        Assert.Contains("rgba(4, 5, 6, 1)", document.StyleSummary.ColorValues);
        Assert.Contains(
            HtmlComputedStyleEngine.Compute(document).Values,
            style => style.GetValue("color") == "rgba(4, 5, 6, 1)");
    }

    [Fact]
    public async Task HtmlConversionDocument_SerializesConcurrentLazySourceAnalysis() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>p{color:rgb(1,2,3)}</style><main><p>Concurrent</p><img src='data:image/png;base64,AQID'></main>");

        await Task.WhenAll(Enumerable.Range(0, 32).Select(index => Task.Run(() => {
            switch (index % 6) {
                case 0: _ = document.LogicalDocument; break;
                case 1: _ = document.StyleSummary; break;
                case 2: _ = document.ResourceManifest; break;
                case 3: _ = document.NormalizedHtml; break;
                case 4: _ = document.HtmlForConversion; break;
                default: _ = document.ResourcePlan; break;
            }
        })));

        Assert.Contains("Concurrent", document.HtmlForConversion, StringComparison.Ordinal);
        Assert.NotEmpty(document.ResourceManifest.Resources);
    }

    [Fact]
    public void HtmlDiagnosticCatalog_CoversEverySharedConversionCode() {
        Assert.All(HtmlConversionDiagnosticCodes.All, code => Assert.True(HtmlDiagnosticCatalog.TryGet(code, out _), code));
    }

    private static HtmlConversionDocumentOptions Options(Action<HtmlConversionLimits> configure) {
        HtmlConversionLimits limits = HtmlConversionLimits.CreateUntrustedProfile();
        configure(limits);
        return new HtmlConversionDocumentOptions { Limits = limits };
    }
}
