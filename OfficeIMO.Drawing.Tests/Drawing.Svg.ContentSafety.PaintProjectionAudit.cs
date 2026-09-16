using System.Text;
using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class SvgContentSafetyPaintProjectionAuditTests {
    [Theory]
    [InlineData("same.svg#label")]
    [InlineData("#missing")]
    public void UnresolvedUseReferenceCannotAuthorizeRemovalOfSourceText(string reference) {
        byte[] svg = Svg(
            "<g display='none'><text id='label' x='10' y='35'>referenced payload</text></g>" +
            "<use href='" + reference + "'/>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "referenced payload");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("outside the bounded native paint projection", finding.Evidence, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("linearGradient", "x2")]
    [InlineData("radialGradient", "r")]
    public void UnsupportedGradientCoordinatesCannotAuthorizeTransparentTextRemoval(string kind, string coordinate) {
        byte[] svg = Svg(
            "<defs><" + kind + " id='paint' gradientUnits='userSpaceOnUse' " + coordinate + "='1em'>" +
            "<stop offset='0' stop-color='black'/><stop offset='1' stop-color='black'/>" +
            "</" + kind + "></defs>" +
            "<text fill='url(#paint)' x='10' y='35'>gradient payload</text>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "gradient payload");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("outside the bounded native paint projection", finding.Evidence, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("fill='CanvasText'")]
    [InlineData("stroke='CanvasText'")]
    [InlineData("color='CanvasText' fill='currentColor'")]
    public void UnsupportedBrowserPaintCannotAuthorizeTransparentTextRemoval(string paint) {
        byte[] svg = Svg(
            "<g fill='none' stroke='none'><text " + paint +
            " x='10' y='35'>system-color payload</text></g>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "system-color payload");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("outside the bounded native paint projection", finding.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void UnsupportedGradientStopColorCannotAuthorizeTransparentTextRemoval() {
        byte[] svg = Svg(
            "<defs><linearGradient id='paint'><stop offset='0' stop-color='CanvasText'/>" +
            "<stop offset='1' stop-color='CanvasText'/></linearGradient></defs>" +
            "<text fill='url(#paint)' x='10' y='35'>stop-color payload</text>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "stop-color payload");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("<pattern id='paint' patternUnits='userSpaceOnUse' width='1em' height='20'><rect width='20' height='20' fill='white'/></pattern>")]
    [InlineData("<pattern id='template' width='20' height='20'><rect width='20' height='20' fill='white'/></pattern><pattern id='paint' href='#template'/>")]
    public void UnsupportedPatternGeometryOrTemplateCannotHideEarlierText(string pattern) {
        byte[] svg = Svg(
            "<text x='10' y='35'>pattern-covered payload</text>" +
            "<defs>" + pattern + "</defs>" +
            "<rect x='0' y='0' width='220' height='60' fill='url(#paint)'/>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "pattern-covered payload");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("outside the bounded native paint projection", finding.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void UnsupportedStrokeDashGeometryMarksPaintProjectionIncomplete() {
        byte[] svg = Svg(
            "<text x='10' y='35'>dash-covered payload</text>" +
            "<path d='M0 25h220' fill='none' stroke='white' stroke-width='20' stroke-dasharray='1em 1em'/>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions).Findings,
            item => item.TextPreview == "dash-covered payload");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("outside the bounded native paint projection", finding.Evidence, StringComparison.Ordinal);
    }

    private static byte[] Svg(string body) => Encoding.UTF8.GetBytes(
        "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='120' viewBox='0 0 220 120'>" + body + "</svg>");
}
