using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave25Tests {
    [Fact]
    public void AuthoredTextIndentKeepsPositionedRegionInSemanticFlow() {
        const string html = "<div style='position:absolute;width:180px;height:60px;text-indent:72pt'>Indented text</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Indented text", projection.RemainingDocument.Body!.TextContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Message.Contains("rich document content", StringComparison.Ordinal));
    }

    [Fact]
    public void StylesheetTextIndentKeepsPositionedRegionInSemanticFlow() {
        const string html = "<style>.indented{text-indent:72pt}</style>"
            + "<div class='indented' style='position:absolute;width:180px;height:60px'>Indented text</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Indented text", projection.RemainingDocument.Body!.TextContent);
    }

    [Fact]
    public void DescendantTextIndentKeepsOwningRegionInSemanticFlow() {
        const string html = "<div style='position:absolute;width:180px;height:60px'>"
            + "<div style='text-indent:72pt'>Indented child</div></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Indented child", projection.RemainingDocument.Body!.TextContent);
    }
}