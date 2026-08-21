using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave26Tests {
    [Theory]
    [InlineData("column-count:2", "column-count=2")]
    [InlineData("column-width:80px", "column-width=80px")]
    [InlineData("columns:2 80px", "columns=2 80px")]
    public void MultiColumnRegionsStayInSemanticFlow(string declaration, string expectedDetail) {
        string html = "<div style='position:absolute;width:220px;height:80px;" + declaration
            + "'>Column content that would otherwise be flattened</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains(expectedDetail, StringComparison.OrdinalIgnoreCase));
    }

    [Theory]
    [InlineData("display:flex;justify-content:center", "justify-content=center")]
    [InlineData("display:flex;align-items:center", "align-items=center")]
    [InlineData("display:grid;place-items:end", "place-items=end")]
    public void AlignedSingleItemFlexAndGridRegionsStayInSemanticFlow(
        string declaration,
        string expectedDetail) {
        string html = "<div style='position:absolute;width:220px;height:80px;" + declaration
            + "'><span>Placed content</span></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains(expectedDetail, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ExplicitSingleGridItemPlacementKeepsOwningRegionInSemanticFlow() {
        const string html = "<div style='position:absolute;display:grid;grid-template-columns:40px 40px;"
            + "width:220px;height:80px'><span style='grid-column:2'>Placed content</span></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("grid-column", StringComparison.OrdinalIgnoreCase));
    }
}
