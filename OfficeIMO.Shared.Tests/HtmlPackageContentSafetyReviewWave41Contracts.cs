using OfficeIMO.ContentSafety;
using OfficeIMO.Html;
using OfficeIMO.Mhtml;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class HtmlPackageContentSafetyReviewWave41ContractTests {
    [Theory]
    [InlineData("background-clip:text")]
    [InlineData("-webkit-background-clip:text")]
    [InlineData("background-cl\\69p:text")]
    [InlineData("background-clip:border-box,text")]
    public void Mhtml_BackgroundPaintedTransparentTextIsReportOnly(string clip) {
        byte[] input = new MhtmlDocument(
            "<html><body><p style='color:transparent;background:linear-gradient(red,blue);" + clip +
            "'>Background-painted text.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TransparentText
            && item.TextPreview.Contains("Background-painted text", StringComparison.Ordinal));
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Throws<InvalidOperationException>(() => MhtmlDocument.RemoveSelectedContent(
            input, new OfficeContentCleanupSelection(new[] { finding.Id })));
    }

    [Theory]
    [InlineData("<p style='font-size:1px;zoom:20'>Scaled text.</p>")]
    [InlineData("<style>p{font-size:1px;zoom:20}</style><p>Scaled text.</p>")]
    [InlineData("<p style='font-size:1px;scale:20'>Scaled text.</p>")]
    [InlineData("<div style='transform:scale(20)'><p style='font-size:1px'>Scaled text.</p></div>")]
    public void Mhtml_ScaledTinyTextIsReportOnly(string body) {
        byte[] input = new MhtmlDocument(
            "<html><body>" + body + "</body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TinyText
            && item.TextPreview.Contains("Scaled text", StringComparison.Ordinal));
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("<style>#secret{display:none}#secret{all:initial}</style>")]
    [InlineData("<style>#secret{display:none}#secret{a\\6c l:initial}</style>")]
    [InlineData("<style>#secret{display:none}</style><p style='all:initial'>x</p>")]
    public void Mhtml_CascadeResetFailsClosed(string styles) {
        byte[] input = new MhtmlDocument(
            "<html><head>" + styles + "</head><body><p id='secret'>Secret text.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        Assert.Throws<InvalidDataException>(() => MhtmlDocument.InspectContentSafety(input));
    }

    [Fact]
    public void Html_OfficeConditionalCommentIsReportOnly() {
        const string html = "<html><body><!--[if mso]><p>Office-visible content.</p><![endif]--></body></html>";
        OfficeContentSafetyFinding finding = Assert.Single(HtmlContentSafety.Inspect(html).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.NonPrimaryContent
            && item.TextPreview.Contains("Office-visible content", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Throws<InvalidOperationException>(() => HtmlContentSafety.RemoveSelected(
            html, new OfficeContentCleanupSelection(new[] { finding.Id })));
    }

    [Fact]
    public void Html_OrdinaryCommentRemainsCleanable() {
        const string html = "<html><body><!--Ordinary concealed comment.--></body></html>";
        OfficeContentSafetyFinding finding = Assert.Single(HtmlContentSafety.Inspect(html).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.NonPrimaryContent
            && item.TextPreview.Contains("Ordinary concealed comment", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.RemoveElement, finding.CleanupCapability);
    }

    [Fact]
    public void Mhtml_NoOpClipAndZoomRemainCleanable() {
        byte[] input = new MhtmlDocument(
            "<html><body><p style='color:transparent;background-clip:border-box;zoom:1'>" +
            "Genuinely concealed text.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.TransparentText
            && item.TextPreview.Contains("Genuinely concealed text", StringComparison.Ordinal));
        Assert.NotEqual(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void Mhtml_CssStringMentionOfResetDoesNotRejectPackage() {
        byte[] input = new MhtmlDocument(
            "<html><head><style>p::before{content:'all:initial'}</style></head>" +
            "<body><p style='display:none'>Hidden text.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Hidden text", StringComparison.Ordinal));
        Assert.NotEqual(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void Html_CascadeResetLeavesComputedConcealmentReportOnly() {
        const string html = "<html><head><style>p{display:none}p{all:initial}</style></head>" +
            "<body><p>Reset-visible text.</p></body></html>";

        OfficeContentSafetyFinding finding = Assert.Single(HtmlContentSafety.Inspect(html).Findings, item =>
            item.TextPreview.Contains("Reset-visible text", StringComparison.Ordinal));
        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }
}
