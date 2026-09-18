using OfficeIMO.ContentSafety;
using OfficeIMO.Mhtml;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class HtmlPackageContentSafetyReviewWave40ContractTests {
    [Fact]
    public void Mhtml_SvgFillPaintKeepsTransparentCssColorReportOnly() {
        byte[] input = new MhtmlDocument(
            "<html><body><svg xmlns='http://www.w3.org/2000/svg'><text style='color:transparent;fill:black'>" +
            "SVG fill-painted text.</text></svg></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TransparentText
            && item.TextPreview.Contains("SVG fill-painted text", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("fullscreen")]
    [InlineData("picture-in-picture")]
    [InlineData("full\\73 creen")]
    public void Mhtml_NativeMediaStateRevealIsReportOnly(string pseudoClass) {
        byte[] input = new MhtmlDocument(
            "<html><head><style>#secret{display:none}video:" + pseudoClass +
            "+#secret{display:block}</style></head><body><video></video><p id='secret'>" +
            "Native media-state reveal.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Native media-state reveal", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("max-width:0;min-width:100px")]
    [InlineData("max-height:0;min-height:100px")]
    public void Mhtml_MinimumSizeConflictKeepsZeroDimensionReportOnly(string dimensions) {
        byte[] input = new MhtmlDocument(
            "<html><body><p style='" + dimensions + ";overflow:hidden'>" +
            "Minimum-constrained visible text.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.ZeroDimension
            && item.TextPreview.Contains("Minimum-constrained visible text", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void Mhtml_TransformScaledTinyTextIsReportOnly() {
        byte[] input = new MhtmlDocument(
            "<html><body><p style='font-size:1px;transform:scale(20)'>Transform-scaled readable text.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TinyText
            && item.TextPreview.Contains("Transform-scaled readable text", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void Mhtml_ScriptRevealableHiddenInputAndValueAreReportOnly() {
        byte[] input = new MhtmlDocument(
            "<html><body><input id='x' type='hidden' value='Script-visible control value'>" +
            "<script>document.getElementById('x').type='text'</script></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyReport report = MhtmlDocument.InspectContentSafety(input);
        OfficeContentSafetyFinding valueFinding = Assert.Single(report.Findings, item =>
            item.Location.EndsWith("/@value", StringComparison.Ordinal)
            && item.TextPreview.Contains("Script-visible control value", StringComparison.Ordinal));
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, valueFinding.CleanupCapability);
    }

    [Fact]
    public void Mhtml_LegacyBackgroundHintKeepsLowContrastReportOnly() {
        byte[] input = new MhtmlDocument(
            "<html><body bgcolor='black'><p style='color:white'>Legacy background-visible text.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.LowContrastText
            && item.TextPreview.Contains("Legacy background-visible text", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }
}
