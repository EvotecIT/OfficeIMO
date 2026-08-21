using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave17Tests {
    [Fact]
    public void ProjectionRenderIntersectsCallerAndOwningDocumentNodeLimits() {
        var limits = HtmlConversionLimits.CreateTrustedProfile();
        limits.MaxHtmlNodes = 64;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<div style='position:absolute;width:120px;height:30px'>Projected</div>" +
            "<p>One</p><p>Two</p><p>Three</p>",
            new HtmlConversionDocumentOptions { Limits = limits });

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlEditableLayoutProjector.Project(
                document,
                new HtmlRenderOptions { MaxHtmlNodes = 2 }));

        Assert.Equal(HtmlRenderDiagnosticCodes.NodeLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxHtmlNodes), exception.LimitSource);
        Assert.Equal(2, exception.Limit);
    }

    [Fact]
    public void ManySiblingSemanticRegionsDoNotSuppressLaterProjectionCandidates() {
        var html = new StringBuilder();
        for (int index = 0; index < 512; index++) {
            html.Append("<div style='position:absolute;width:10px;height:10px'><strong>Rich</strong></div>");
        }
        html.Append("<div style='position:absolute;width:120px;height:30px'>Projected</div>");

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html.ToString()));

        Assert.Single(projection.Regions);
        Assert.Equal(512, projection.Diagnostics.Count(diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified));
    }

    [Theory]
    [InlineData("transparent")]
    [InlineData("rgba(20,40,60,0.5)")]
    public void AlphaBackgroundRegionsStayInDiagnosedSemanticFlow(string backgroundColor) {
        string html = "<div style='position:absolute;width:120px;height:30px;background-color:" +
            backgroundColor + "'>Alpha fill</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Alpha fill", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("background-color=", StringComparison.Ordinal));
    }

    [Fact]
    public void ShortConfiguredWordPageKeepsContinuousRegionsInSemanticFlow() {
        const string html = "<div style='position:absolute;top:500px;width:120px;height:30px'>Page owned</div>";
        var options = new HtmlToWordOptions {
            DefaultPageSize = WordPageSize.A6,
            DefaultOrientation = OfficePageOrientation.Landscape
        };

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
        using WordDocument word = result.Value;

        Assert.Empty(word.TextBoxes);
        Assert.NotEmpty(word.Find("Page owned", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("maximumPageHeight", StringComparison.Ordinal));
    }

    [Fact]
    public void WordAnchorSizesClampToDrawingMlPositiveCoordinateRange() {
        long actual = HtmlToWordConverter.ToBoundedAnchorSize(
            HtmlToWordConverter.MaximumDrawingExtent + 1D,
            1D,
            out bool simplified);

        Assert.True(simplified);
        Assert.Equal(HtmlToWordConverter.MaximumDrawingExtent, actual);
    }
}