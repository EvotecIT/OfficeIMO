using System.Text;
using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class SvgContentSafetyStructuralSyntaxTests {
    [Theory]
    [InlineData("<rect/>")]
    [InlineData("<rect width='20'/>")]
    [InlineData("<rect width='-1' height='20'/>")]
    [InlineData("<circle/>")]
    [InlineData("<ellipse/>")]
    public void MissingClipDimensionsLeaveNoVisibleGeometry(string geometry) {
        byte[] svg = Svg("<defs><clipPath id='empty'>" + geometry + "</clipPath></defs>" +
            "<text clip-path='url(#empty)' x='10' y='35'>empty clip payload</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions).Findings,
            item => item.TextPreview == "empty clip payload");

        Assert.Equal(OfficeContentConcealmentKind.ClippedContent, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.RemoveText, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("<ellipse rx='auto' ry='auto'/>")]
    [InlineData("<line x1='0' y1='0' x2='20' y2='20'/>")]
    [InlineData("<path d='M0 0'/>")]
    public void EmptyClipOutsideExactNativeGeometryIsStillDetected(string geometry) {
        byte[] svg = Svg("<defs><clipPath id='empty'>" + geometry + "</clipPath></defs>" +
            "<text clip-path='url(#empty)' x='10' y='35'>empty clip payload</text>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "empty clip payload");

        Assert.Equal(OfficeContentConcealmentKind.ClippedContent, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void StyledClipDimensionsDoNotBecomeAnEmptyStructuralClip() {
        byte[] svg = Svg("<defs><clipPath id='styled'><rect style='width:100px;height:100px'/></clipPath></defs>" +
            "<text clip-path='url(#styled)' x='10' y='35'>styled clip payload</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(
            svg, readerOptions: readerOptions));
    }

    [Fact]
    public void OneEllipseRadiusCanSupplyTheOtherWithoutAnEmptyClipFinding() {
        byte[] svg = Svg("<defs><clipPath id='round'><ellipse rx='20'/></clipPath></defs>" +
            "<text clip-path='url(#round)' x='10' y='35'>round clip payload</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "round clip payload" &&
            item.Kind == OfficeContentConcealmentKind.ClippedContent &&
            item.CleanupCapability == OfficeContentCleanupCapability.RemoveText);
    }

    [Theory]
    [InlineData("<ellipse rx='-1' ry='10'/>")]
    [InlineData("<ellipse rx='10' ry='-1'/>")]
    public void InvalidNegativeEllipseRadiusFallsBackToTheOtherRadius(string geometry) {
        byte[] svg = Svg("<defs><clipPath id='round'>" + geometry + "</clipPath></defs>" +
            "<text clip-path='url(#round)' x='10' y='35'>negative radius payload</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "negative radius payload" &&
            item.Kind == OfficeContentConcealmentKind.ClippedContent &&
            item.CleanupCapability == OfficeContentCleanupCapability.RemoveText);
    }

    [Theory]
    [InlineData("opacity='0 %'")]
    [InlineData("fill-opacity='0 %'")]
    [InlineData("stroke-opacity='0 %'")]
    [InlineData("opacity='0\t%'")]
    public void SeparatedPercentageCannotAuthorizeTransparentTextCleanup(string attribute) {
        byte[] svg = Svg("<text " + attribute + " x='10' y='35'>visible percentage payload</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible percentage payload" &&
            item.CleanupCapability == OfficeContentCleanupCapability.RemoveText);
    }

    [Fact]
    public void SeparatedPixelUnitCannotAuthorizeTinyTextCleanup() {
        byte[] svg = Svg("<text font-size='0 px' x='10' y='35'>visible font payload</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible font payload" &&
            item.CleanupCapability == OfficeContentCleanupCapability.RemoveText);
    }

    [Theory]
    [InlineData("opacity='0.'")]
    [InlineData("opacity='0.px'")]
    [InlineData("fill-opacity='0.'")]
    public void InvalidNumericTokenCannotAuthorizeTransparentTextCleanup(string attribute) {
        byte[] svg = Svg("<text " + attribute + " x='10' y='35'>visible numeric payload</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible numeric payload" &&
            item.CleanupCapability == OfficeContentCleanupCapability.RemoveText);
    }

    [Theory]
    [InlineData("opacity:0.")]
    [InlineData("opacity:0.%")]
    [InlineData("fill-opacity:0.")]
    public void InvalidNumericCssDeclarationFailsClosed(string declaration) {
        byte[] svg = Svg("<text style='" + declaration + "' x='10' y='35'>visible CSS numeric payload</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Theory]
    [InlineData("1.px")]
    [InlineData("1e+px")]
    public void InvalidFontNumericTokenCannotAuthorizeTinyTextCleanup(string size) {
        byte[] svg = Svg("<text font-size='" + size + "' x='10' y='35'>visible font payload</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible font payload" &&
            item.CleanupCapability == OfficeContentCleanupCapability.RemoveText);
    }

    [Theory]
    [InlineData("10 px")]
    [InlineData("10.px")]
    [InlineData("1e+px")]
    public void InvalidRootWidthCannotProjectAViewportForCleanup(string width) {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='" + width +
            "' height='120' viewBox='0 0 220 120'><text font-size='1' x='10' y='35'>visible viewport payload</text></svg>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Theory]
    [InlineData("opacity:0 %")]
    [InlineData("fill-opacity:0 %")]
    public void SeparatedPercentageInInlineStyleFailsClosed(string declaration) {
        byte[] svg = Svg("<text style='" + declaration + "' x='10' y='35'>invalid style payload</text>");

        Assert.Throws<InvalidDataException>(() => OfficeSvgDrawingReader.InspectContentSafety(svg));
    }

    [Fact]
    public void SeparatedBaselineShiftUnitCannotAuthorizeOffCanvasCleanup() {
        byte[] svg = Svg("<text baseline-shift='1000 %' x='10' y='35'>visible baseline payload</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible baseline payload" &&
            item.Kind == OfficeContentConcealmentKind.OffCanvas &&
            item.CleanupCapability == OfficeContentCleanupCapability.RemoveText);
    }

    [Theory]
    [InlineData("2.5", true)]
    [InlineData("3", false)]
    public void TinyFontThresholdConvertsCssPixelsToPoints(string fontSize, bool isTiny) {
        byte[] svg = Svg("<text font-size='" + fontSize + "' x='10' y='35'>font threshold payload</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions);

        Assert.Equal(isTiny, report.Findings.Any(item => item.TextPreview == "font threshold payload" &&
            item.Kind == OfficeContentConcealmentKind.TinyText));
    }

    [Fact]
    public void PaintEscapingLaterNestedViewportReportsEarlierTextAsProjectionIncomplete() {
        byte[] svg = Svg(
            "<text x='120' y='35'>overflow-covered payload</text>" +
            "<svg width='100' height='60' overflow='visible'>" +
            "<rect x='110' y='0' width='110' height='60' fill='white'/></svg>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "overflow-covered payload");

        Assert.Equal(OfficeContentConcealmentKind.NonPrimaryContent, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("outside the bounded native paint projection", finding.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void ReferencedSymbolOverflowCannotHideEarlierTextFromInspection() {
        byte[] svg = Svg(
            "<text x='120' y='35'>symbol-covered payload</text>" +
            "<defs><symbol id='cover' viewBox='0 0 100 60' overflow='visible'>" +
            "<rect x='110' y='0' width='110' height='60' fill='white'/></symbol></defs>" +
            "<use href='#cover' width='100' height='60'/>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "symbol-covered payload");

        Assert.Equal(OfficeContentConcealmentKind.NonPrimaryContent, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("outside the bounded native paint projection", finding.Evidence, StringComparison.Ordinal);
    }

    [Fact]
    public void UseShadowTreeInheritanceCannotHideEarlierTextFromInspection() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='400' height='120' viewBox='0 0 400 120'>" +
            "<text x='245' y='35'>use-covered payload</text>" +
            "<defs><rect id='cover' transform='inherit' x='0' y='0' width='110' height='60' fill='white'/></defs>" +
            "<use href='#cover' transform='translate(120,0)'/></svg>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "use-covered payload");

        Assert.Equal(OfficeContentConcealmentKind.NonPrimaryContent, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("outside the bounded native paint projection", finding.Evidence, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("none")]
    [InlineData("initial")]
    [InlineData("unset")]
    public void IdentityTransformKeywordsRemainWithinTheSupportedSubset(string transform) {
        byte[] svg = Svg("<text transform='" + transform + "' x='10' y='35'>visible transform payload</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg);

        Assert.DoesNotContain(report.Findings, item => item.TextPreview == "visible transform payload" &&
            item.Kind == OfficeContentConcealmentKind.NonPrimaryContent);
    }

    [Fact]
    public void UnsupportedBaselineShiftLengthCannotSilentlyHideAnOcclusionCandidate() {
        byte[] svg = Svg(
            "<text x='10' y='100' baseline-shift='72pt'>baseline-shift payload</text>" +
            "<rect x='0' y='0' width='220' height='55' fill='white'/>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "baseline-shift payload");

        Assert.Equal(OfficeContentConcealmentKind.NonPrimaryContent, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("baseline-shift", finding.Evidence, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UnsupportedLineHeightLengthAlsoDowngradesBaselineGeometry() {
        byte[] svg = Svg(
            "<text x='10' y='100' baseline-shift='50%' line-height='72pt'>line-height payload</text>" +
            "<rect x='0' y='0' width='220' height='55' fill='white'/>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "line-height payload");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("line-height", finding.Evidence, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("4ch")]
    [InlineData("4ex")]
    [InlineData("50%")]
    [InlineData("super")]
    [InlineData("sub")]
    public void FontMetricDependentBaselineShiftIsReportOnly(string baselineShift) {
        byte[] svg = Svg(
            "<text x='10' y='100' font-size='20'><tspan baseline-shift='" + baselineShift +
            "'>metric-shift payload</tspan></text>");

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg).Findings,
            item => item.TextPreview == "metric-shift payload");

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains("baseline-shift", finding.Evidence, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SupportedBaselineShiftStillAllowsExactStructuralCleanup() {
        byte[] svg = Svg("<text x='10' y='35' baseline-shift='8px' opacity='0'>hidden shifted payload</text>");
        var readerOptions = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualComparisons = 0 };

        OfficeContentSafetyFinding finding = Assert.Single(
            OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: readerOptions).Findings,
            item => item.TextPreview == "hidden shifted payload");

        Assert.Equal(OfficeContentConcealmentKind.TransparentText, finding.Kind);
        Assert.Equal(OfficeContentCleanupCapability.RemoveText, finding.CleanupCapability);
    }

    private static byte[] Svg(string body) => Encoding.UTF8.GetBytes(
        "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='120' viewBox='0 0 220 120'>" + body + "</svg>");
}
