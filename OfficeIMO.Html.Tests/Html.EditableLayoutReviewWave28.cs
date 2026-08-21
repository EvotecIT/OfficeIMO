using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave28Tests {
    [Theory]
    [InlineData("<progress value='40' max='100'></progress>")]
    [InlineData("<meter value='0.82' max='1'></meter>")]
    public void ValueControlsStayInDiagnosedSemanticFlow(string control) {
        string html = "<div style='position:absolute;width:220px;height:40px'>" + control + "</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "semanticContent=true");
    }

    [Theory]
    [InlineData("<body style='color:red'><div style='position:absolute;width:220px;height:40px;color:red'>Text</div></body>")]
    [InlineData("<style>body{color:red}.region{position:absolute;width:220px;height:40px;color:red}</style><div class='region'>Text</div>")]
    public void ExplicitTypographyMatchingParentStaysInSemanticFlow(string html) {
        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "semanticContent=true");
    }

    [Fact]
    public void ExplicitDescendantPaintMatchingParentStaysInSemanticFlow() {
        const string html = "<div style='position:absolute;width:220px;height:40px;background-color:red'>"
            + "<span style='background-color:red'>Text</span></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "semanticContent=true");
    }
}
