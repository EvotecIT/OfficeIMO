using OfficeIMO.ContentSafety;
using OfficeIMO.Mhtml;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class HtmlPackageContentSafetyReviewWave39ContractTests {
    [Fact]
    public void Mhtml_OffCanvasAncestorWithCounterPositionedDescendantIsReportOnly() {
        byte[] input = new MhtmlDocument(
            "<html><body><div style='position:absolute;left:-2000pt'>Off-canvas parent." +
            "<span style='position:absolute;left:2000pt'>Counter-positioned visible text.</span>" +
            "</div></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.OffCanvas
            && item.TextPreview.Contains("Counter-positioned visible text", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Throws<InvalidOperationException>(() => MhtmlDocument.RemoveSelectedContent(
            input,
            new OfficeContentCleanupSelection(new[] { finding.Id })));
    }

    [Theory]
    [InlineData("text-shadow:0 0 0 black")]
    [InlineData("te\\78 t-shadow:0 0 0 black")]
    public void Mhtml_VisibleTextShadowMakesTransparentTextReportOnly(string shadowDeclaration) {
        byte[] input = new MhtmlDocument(
            "<html><body><p style='color:transparent;" + shadowDeclaration +
            "'>Shadow-painted visible text.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TransparentText
            && item.TextPreview.Contains("Shadow-painted visible text", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }
}
