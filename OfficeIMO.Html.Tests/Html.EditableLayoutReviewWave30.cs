using OfficeIMO.Html;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave30Tests {
    [Fact]
    public void PreparedProjectionIntersectsCallerAndOwningNodeLimits() {
        var limits = HtmlConversionLimits.CreateTrustedProfile();
        limits.MaxHtmlNodes = 64;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<div style='position:absolute;width:120px;height:30px'>Projected</div>"
            + "<p>One</p><p>Two</p><p>Three</p>",
            new HtmlConversionDocumentOptions { Limits = limits });

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlEditableLayoutProjector.Project(document, new HtmlRenderOptions { MaxHtmlNodes = 2 }));

        Assert.Equal(HtmlRenderDiagnosticCodes.NodeLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxHtmlNodes), exception.LimitSource);
        Assert.Equal(2, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void SoftHyphenationDoesNotBecomeEditablePaintedText() {
        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(
                "<div style='position:absolute;width:50px;height:80px'>co\u00ADoperate</div>"));

        Assert.True(projection.Regions.Count == 1,
            string.Join(" | ", projection.Diagnostics.Select(diagnostic => diagnostic.Code + ":" + diagnostic.Detail)));
        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);
        Assert.Equal("cooperate", region.SourceText);
        Assert.DoesNotContain("-", region.SourceText, StringComparison.Ordinal);
    }

    [Fact]
    public void AutomaticHyphenationDoesNotBecomeEditablePaintedText() {
        var options = new HtmlRenderOptions {
            TextHyphenationCallback = token => token == "typography" ? new[] { 4 } : Array.Empty<int>()
        };

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(
                "<div style='position:absolute;width:50px;height:80px;hyphens:auto'>typography</div>"),
            options);

        Assert.True(projection.Regions.Count == 1,
            string.Join(" | ", projection.Diagnostics.Select(diagnostic => diagnostic.Code + ":" + diagnostic.Detail)));
        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);
        Assert.Equal("typography", region.SourceText);
        Assert.DoesNotContain("-", region.SourceText, StringComparison.Ordinal);
    }
}