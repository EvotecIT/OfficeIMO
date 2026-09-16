using System.Text;
using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class SvgContentSafetyCssBoundaryTests {
    [Fact]
    public void InvalidXmlSpaceCasingFailsClosed() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xml:space='PRESERVE' width='220' height='120' viewBox='0 0 220 120'>" +
            "<rect width='220' height='120' fill='white'/>" +
            "<text font-family='OfficeIMO Shaping Test' font-size='20' x='10' y='35'>      A</text>" +
            "<rect width='40' height='60' fill='white'/></svg>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void InvalidNestedCustomPropertyUsesTheCurrentVarFallback() {
        byte[] svg = Svg(
            "<text style='--visibility:var(--missing);display:var(--visibility,none)' x='10' y='35'>nested invalid custom fallback</text>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "nested invalid custom fallback");

        Assert.Equal(OfficeContentConcealmentKind.HiddenByProperty, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.RemoveText, finding.CleanupCapability);
    }

    [Fact]
    public void NonSvgWhitespaceInViewBoxFailsClosed() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='200' height='100' viewBox='0\u00A00\u00A08000\u00A0100'>" +
            "<text font-size='50' x='10' y='60'>visible text</text></svg>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Theory]
    [InlineData("media='\u00A0all'")]
    [InlineData("type='\u00A0text/css'")]
    [InlineData("title=' '")]
    public void StylesheetControlsOutsideTheBoundedGrammarFailClosed(string attributes) {
        byte[] svg = Svg(
            $"<style {attributes}>text{{display:none}}</style>" +
            "<text x='10' y='35'>browser visible text</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void UnsupportedEmbeddedSvgImageMakesFollowingPaintIncomplete() {
        string image = Convert.ToBase64String(Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='120'><rect width='220' height='120' fill='white'/></svg>"));
        byte[] svg = Svg(
            "<text x='10' y='35'>embedded image payload</text>" +
            $"<image href='data:image/svg+xml;base64,{image}' width='220' height='120'/>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "embedded image payload");

        Assert.Equal(OfficeContentConcealmentKind.NonPrimaryContent, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void UnsupportedEmbeddedSvgImageBeforeTextMakesUnicodeCleanupReportOnly() {
        string image = Convert.ToBase64String(Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='120'><rect width='220' height='120' fill='white'/></svg>"));
        byte[] svg = Svg(
            $"<image href='data:image/svg+xml;base64,{image}' width='220' height='120'/>" +
            "<text x='10' y='35'>pay​load</text>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(
                svg,
                new OfficeContentSafetyOptions { IncludeNonPrimaryContent = false }).Findings,
            item => item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode && item.TextPreview == "\\u200B");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void IntrinsicSizeRasterImageMakesUnicodeCleanupReportOnly() {
        const string gif = "R0lGODlhAQABAJAAAAAAAP///ywAAAAAAQABAAACAkwBADs=";
        byte[] svg = Svg(
            $"<image href='data:image/gif;base64,{gif}'/>" +
            "<text x='10' y='35'>pay​load</text>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(
                svg,
                new OfficeContentSafetyOptions { IncludeNonPrimaryContent = false }).Findings,
            item => item.Kind == OfficeContentConcealmentKind.NonPrintingUnicode && item.TextPreview == "\\u200B");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void UnsupportedPaintProducedByVarFailsClosed() {
        byte[] svg = Svg(
            "<g fill='none'><text style='--paint:url(#missing) red;fill:var(--paint)' x='10' y='35'>browser visible paint</text></g>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void CaseMismatchedPreserveAspectRatioFailsClosed() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='200' height='100' viewBox='0 0 100 100' preserveAspectRatio='xMidYMid SLICE'>" +
            "<rect width='100' height='100' fill='white'/><text x='10' y='10'>visible text</text></svg>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Theory]
    [InlineData("none")]
    [InlineData("none meet")]
    [InlineData("none slice")]
    public void PreserveAspectRatioNoneAcceptsOptionalModeToken(string value) {
        byte[] svg = Encoding.UTF8.GetBytes(
            $"<svg xmlns='http://www.w3.org/2000/svg' width='200' height='100' viewBox='0 0 100 100' preserveAspectRatio='{value}'>" +
            "<rect width='100' height='100' fill='white'/><text x='10' y='20'>visible text</text></svg>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Empty(report.Findings);
    }

    [Fact]
    public void ViewportAttributesOnTextAreIgnoredByViewportGrammar() {
        byte[] svg = Svg(
            "<text viewBox='0 0 100 100' preserveAspectRatio='xMidYMid SLICE' x='10' y='35'>visible text</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.Empty(report.Findings);
    }

    [Fact]
    public void UnsupportedForeignObjectPaintMakesProjectionReportOnly() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xhtml='http://www.w3.org/1999/xhtml' width='220' height='120' viewBox='0 0 220 120'>" +
            "<text x='10' y='35'>foreign object payload</text>" +
            "<foreignObject x='0' y='0' width='220' height='120'><xhtml:div style='width:220px;height:120px;background:white'/></foreignObject></svg>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "foreign object payload");

        Assert.Equal(OfficeContentConcealmentKind.NonPrimaryContent, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("incomplete", finding.Evidence, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void StyledStrokeWidthMakesTinyTextCleanupReportOnly() {
        byte[] svg = Svg(
            "<text style='font-size:1;fill:none;stroke:black;stroke-width:20' x='10' y='35'>styled outline payload</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions).Findings,
            item => item.TextPreview == "styled outline payload" && item.Kind == OfficeContentConcealmentKind.TinyText);

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("stroke-width", finding.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void FontSizeAdjustMakesTinyTextCleanupReportOnly() {
        byte[] svg = Svg(
            "<text font-size='1' font-size-adjust='100' x='10' y='35'>adjusted tiny payload</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions).Findings,
            item => item.TextPreview == "adjusted tiny payload" && item.Kind == OfficeContentConcealmentKind.TinyText);

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("font-size-adjust", finding.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void RepeatedViewBoxCommasFailClosed() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='200' height='100' viewBox='0,,0,,8000,,100'>" +
            "<text font-size='50' x='10' y='60'>visible text</text></svg>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void PolygonPointsAcceptImplicitNegativeCoordinates() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='120' viewBox='0 0 220 120'>" +
            "<polygon points='100-0 220,0 220,120 100,120' fill='black'/></svg>");

        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Single(drawing!.Shapes);
    }

    [Fact]
    public void EstimatedFontMetricsMakeVisualCleanupReportOnly() {
        byte[] svg = Svg(
            "<rect width='220' height='120' fill='white'/>" +
            "<text font-family='DefinitelyMissingOfficeImoFont' font-size='20' fill='black' x='10' y='50'>estimated visual payload</text>" +
            "<rect x='0' y='20' width='220' height='45' fill='white'/>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "estimated visual payload");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("font metrics", finding.Evidence, StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] Svg(string body) => Encoding.UTF8.GetBytes(
        "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='120' viewBox='0 0 220 120'>" + body + "</svg>");
}
