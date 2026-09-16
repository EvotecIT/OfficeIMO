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

    private static byte[] Svg(string body) => Encoding.UTF8.GetBytes(
        "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='120' viewBox='0 0 220 120'>" + body + "</svg>");
}
