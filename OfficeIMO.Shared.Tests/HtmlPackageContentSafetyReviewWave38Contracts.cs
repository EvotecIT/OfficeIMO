using OfficeIMO.ContentSafety;
using OfficeIMO.Email;
using OfficeIMO.Mhtml;
using System.Text;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class HtmlPackageContentSafetyReviewWave38ContractTests {
    [Theory]
    [InlineData("<object data='data:text/html,%3Cp%20hidden%3ENested%3C/p%3E'></object>")]
    [InlineData("<embed src='nested.html' type='text/html'>")]
    [InlineData("<portal src='nested.html'></portal>")]
    [InlineData("<fencedframe src='nested.html'></fencedframe>")]
    public void Mhtml_RejectsAdditionalNestedRenderingContexts(string embeddedDocument) {
        byte[] input = new MhtmlDocument(
            "<html><body>" + embeddedDocument + "</body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        Assert.Throws<InvalidDataException>(() => MhtmlDocument.InspectContentSafety(input));
    }

    [Theory]
    [InlineData("animate")]
    [InlineData("animateColor")]
    [InlineData("animateMotion")]
    [InlineData("animateTransform")]
    [InlineData("discard")]
    [InlineData("set")]
    public void Mhtml_SvgDeclarativeAnimationMakesComputedConcealmentReportOnly(string animationElement) {
        byte[] input = new MhtmlDocument(
            "<html><body><svg xmlns='http://www.w3.org/2000/svg'>" +
            "<text style='opacity:0'>Declaratively revealed." +
            "<" + animationElement + " attributeName='opacity' to='1' dur='1s' /></text></svg></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Declaratively revealed", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("until-found")]
    [InlineData("UNTIL-FOUND")]
    public void Mhtml_HiddenUntilFoundContentIsReportOnly(string hiddenValue) {
        byte[] input = new MhtmlDocument(
            "<html><body><section hidden='" + hiddenValue + "'>Search-revealable content.</section></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Search-revealable content", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Throws<InvalidOperationException>(() => MhtmlDocument.RemoveSelectedContent(
            input,
            new OfficeContentCleanupSelection(new[] { finding.Id })));
    }

    [Fact]
    public void Mhtml_CleanupSerializationUsesTheReinspectionInputLimit() {
        byte[] input = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/related; type=\"text/html\"; boundary=\"b\"\r\n\r\n" +
            "--b\r\nContent-Type: text/html\r\nContent-Location: https://example.test/index.html\r\n\r\n" +
            "<p hidden>x</p>\r\n--b--\r\n");
        var inspection = new OfficeContentSafetyOptions {
            MaxInputBytes = input.LongLength,
            MaxExpandedPackageBytes = input.LongLength * 8
        };
        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input, inspection).Findings, item =>
            item.TextPreview.Contains("x", StringComparison.Ordinal));

        EmailLimitExceededException exception = Assert.Throws<EmailLimitExceededException>(() =>
            MhtmlDocument.RemoveSelectedContent(
                input,
                new OfficeContentCleanupSelection(new[] { finding.Id }),
                new OfficeContentCleanupOptions { Inspection = inspection }));

        Assert.Equal(nameof(EmailWriterOptions.MaxOutputBytes), exception.LimitName);
        Assert.Equal(input.LongLength, exception.MaximumValue);
    }
}
