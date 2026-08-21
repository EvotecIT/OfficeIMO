using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutReviewWave32Tests {
    [Theory]
    [InlineData("<span style='display:inline-block;width:100px'>A</span><span>B</span>")]
    [InlineData("<span style='display:inline-block;min-width:100px'>A</span><span>B</span>")]
    [InlineData("<span style='display:inline-block;max-width:100px'>A</span><span>B</span>")]
    public void SizedInlineFormattingContextsStayInDiagnosedSemanticFlow(string children) {
        string html = "<div style='position:absolute;width:180px;height:50px'>" + children + "</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Equal("AB", projection.RemainingDocument.Body!.TextContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("descendant:", StringComparison.Ordinal)
            && diagnostic.Detail.Contains("semanticFlow=true", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DescendantBackgroundImagesStayInDiagnosedSemanticFlow(bool useStylesheet) {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string child = useStylesheet
            ? "<style>.painted-child{background-image:url('" + image + "')}</style>"
                + "<span class='painted-child'>Painted</span>"
            : "<span style=\"background-image:url('" + image + "')\">Painted</span>";
        string html = "<div style='position:absolute;width:180px;height:50px'>" + child + "</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Painted", projection.RemainingDocument.Body!.TextContent, StringComparison.Ordinal);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("descendant:background-image=", StringComparison.Ordinal)
            && diagnostic.Detail.Contains("semanticFlow=true", StringComparison.Ordinal));
    }
}