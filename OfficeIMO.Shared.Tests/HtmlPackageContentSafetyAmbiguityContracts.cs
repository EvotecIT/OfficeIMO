using OfficeIMO.ContentSafety;
using OfficeIMO.Mhtml;
using System.Text;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class HtmlPackageContentSafetyAmbiguityContractTests {
    [Fact]
    public void Mhtml_HiddenAttributeOverrideRemainsVisible() {
        byte[] input = new MhtmlDocument(
            "<html><head><style>[hidden]{display:block}</style></head>" +
            "<body><p hidden>Visible override.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyReport report = MhtmlDocument.InspectContentSafety(input);

        Assert.DoesNotContain(report.Findings, finding =>
            finding.TextPreview.Contains("Visible override", StringComparison.Ordinal));
    }

    [Fact]
    public void Mhtml_ResponsiveHiddenAttributeOverrideIsReportOnly() {
        byte[] input = new MhtmlDocument(
            "<html><head><style>@media (max-width:600px){[hidden]{display:block}}</style></head>" +
            "<body><p hidden>Responsive override.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Responsive override", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("@keyframes reveal{to{opacity:1}}")]
    [InlineData("@-webkit-keyframes reveal{to{opacity:1}}")]
    [InlineData("@\\6b eyframes reveal{to{opacity:1}}")]
    public void Mhtml_AnimatedConcealmentIsReportOnly(string keyframes) {
        byte[] input = new MhtmlDocument(
            "<html><head><style>" + keyframes +
            ".animated{opacity:0;animation:reveal 1s forwards}</style></head>" +
            "<body><p class='animated'>Animated reveal.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Animated reveal", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("@container")]
    [InlineData("@\\63 ontainer")]
    public void Mhtml_ContainerQueryConcealmentIsReportOnly(string containerAtRule) {
        byte[] input = new MhtmlDocument(
            "<html><head><style>" + containerAtRule +
            " (min-width:400px){.responsive{display:none}}</style></head>" +
            "<body><div style='container-type:inline-size;width:50vw'>" +
            "<p class='responsive'>Container responsive.</p></div></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Container responsive", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("é")]
    [InlineData("boundary[")]
    [InlineData("aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa")]
    public void Mhtml_RejectsInvalidMultipartBoundarySyntax(string boundary) {
        byte[] input = BuildMhtmlWithBoundary(boundary, boundary == "é" ? "?" : boundary);
        using (var stream = new MemoryStream(input, writable: false)) {
            MhtmlDocument document = MhtmlDocument.Load(stream);
            Assert.Contains(document.MimeDiagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_MIME_BOUNDARY_INVALID");
        }

        Assert.Throws<InvalidDataException>(() => MhtmlDocument.InspectContentSafety(input));
    }

    [Fact]
    public void Mhtml_RejectsUnmodeledPreferredAlternative() {
        byte[] input = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/related; boundary=outer\r\n\r\n" +
            "--outer\r\n" +
            "Content-Type: multipart/alternative; boundary=inner\r\n\r\n" +
            "--inner\r\n" +
            "Content-Type: text/html; charset=utf-8\r\n\r\n" +
            "<html><body><p style='display:none'>Earlier HTML.</p></body></html>\r\n" +
            "--inner\r\n" +
            "Content-Type: application/xhtml+xml\r\n\r\n" +
            "<html xmlns='http://www.w3.org/1999/xhtml'><body><p>Preferred XHTML.</p></body></html>\r\n" +
            "--inner--\r\n" +
            "--outer--\r\n");

        Assert.Throws<InvalidDataException>(() => MhtmlDocument.InspectContentSafety(input));
    }

    [Theory]
    [InlineData("Content-Security-Policy")]
    [InlineData("X-Content-Security-Policy")]
    [InlineData("X-WebKit-CSP")]
    public void Mhtml_RejectsActiveRootContentSecurityPolicyHeaders(string headerName) {
        byte[] input = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/related; boundary=outer\r\n\r\n" +
            "--outer\r\n" +
            "Content-Type: text/html; charset=utf-8\r\n" +
            headerName + ": style-src 'none'\r\n\r\n" +
            "<html><head><style>.concealed{display:none}</style></head>" +
            "<body><p class='concealed'>CSP-visible content.</p></body></html>\r\n" +
            "--outer--\r\n");

        Assert.Throws<InvalidDataException>(() => MhtmlDocument.InspectContentSafety(input));
    }

    [Fact]
    public void Mhtml_RejectsDuplicateSnapshotContentLocationHeaders() {
        byte[] input = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\n" +
            "Snapshot-Content-Location: https://first.example/index.html\r\n" +
            "Snapshot-Content-Location: https://second.example/index.html\r\n" +
            "Content-Type: multipart/related; boundary=outer\r\n\r\n" +
            "--outer\r\n" +
            "Content-Type: text/html; charset=utf-8\r\n\r\n" +
            "<html><body><p style='display:none'>Ambiguous snapshot base.</p></body></html>\r\n" +
            "--outer--\r\n");

        using (var stream = new MemoryStream(input, writable: false)) {
            MhtmlDocument document = MhtmlDocument.Load(stream);
            Assert.Contains(document.MimeDiagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_MIME_SINGLETON_HEADER_DUPLICATE");
        }
        Assert.Throws<InvalidDataException>(() => MhtmlDocument.InspectContentSafety(input));
    }

    private static byte[] BuildMhtmlWithBoundary(string declaredBoundary, string wireBoundary) => Encoding.UTF8.GetBytes(
        "MIME-Version: 1.0\r\n" +
        "Content-Type: multipart/related; boundary=\"" + declaredBoundary + "\"\r\n\r\n" +
        "--" + wireBoundary + "\r\n" +
        "Content-Type: text/html; charset=utf-8\r\n\r\n" +
        "<html><body><p style='display:none'>Invalid boundary.</p></body></html>\r\n" +
        "--" + wireBoundary + "--\r\n");
}
