using OfficeIMO.ContentSafety;
using OfficeIMO.Email;
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
    [InlineData(":target")]
    [InlineData(":hover")]
    [InlineData(":focus")]
    [InlineData(":checked")]
    [InlineData(":indeterminate")]
    [InlineData(":enabled")]
    [InlineData(":disabled")]
    [InlineData(":\\74 arget")]
    public void Mhtml_StateDependentConcealmentIsReportOnly(string stateSelector) {
        byte[] input = new MhtmlDocument(
            "<html><head><style>#secret{display:none}#secret" + stateSelector +
            "{display:block}</style></head><body><p id='secret'>State reveal.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("State reveal", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("<script>document.getElementById('secret').style.display='block'</script>")]
    [InlineData("<button onclick=\"document.getElementById('secret').style.display='block'\">Reveal</button>")]
    [InlineData("<svg onload=\"document.getElementById('secret').style.display='block'\"></svg>")]
    public void Mhtml_ScriptableConcealmentIsReportOnly(string scriptingMarkup) {
        byte[] input = new MhtmlDocument(
            "<html><body><p id='secret' style='display:none'>Script reveal.</p>" + scriptingMarkup + "</body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Script reveal", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void Mhtml_NativeMutableAttributeConcealmentIsReportOnly() {
        byte[] input = new MhtmlDocument(
            "<html><head><style>details:not([open]) #secret{display:none}</style></head>" +
            "<body><details><summary>Reveal</summary><p id='secret'>Native open reveal.</p></details></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Native open reveal", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void Mhtml_DeclarativeShadowRootContentIsReportOnly() {
        byte[] input = new MhtmlDocument(
            "<html><body><div><template shadowrootmode='open'><p>Shadow-visible content.</p></template></div></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Shadow-visible content", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("srcdoc='&lt;p hidden&gt;Nested concealed.&lt;/p&gt;'")]
    [InlineData("src='frame.html'")]
    [InlineData("src='data:text/html,%3Cp%20hidden%3ENested%3C/p%3E'")]
    public void Mhtml_RejectsNestedBrowsingContexts(string frameAttribute) {
        byte[] input = new MhtmlDocument(
            "<html><body><iframe " + frameAttribute + "></iframe></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        Assert.Throws<InvalidDataException>(() => MhtmlDocument.InspectContentSafety(input));
    }

    [Fact]
    public void Mhtml_RejectsAutomaticRefreshNavigation() {
        byte[] input = new MhtmlDocument(
            "<html><head><meta http-equiv='refresh' content='0; url=next.html'></head>" +
            "<body><p>Inspected root.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        Assert.Throws<InvalidDataException>(() => MhtmlDocument.InspectContentSafety(input));
    }

    [Theory]
    [InlineData("mix-blend-mode:difference")]
    [InlineData("mix-\\62 lend-mode:difference")]
    [InlineData("filter:invert(1)")]
    [InlineData("-webkit-filter:invert(1)")]
    [InlineData("backdrop-filter:invert(1)")]
    public void Mhtml_CompositedLowContrastIsReportOnly(string compositingStyle) {
        byte[] input = new MhtmlDocument(
            "<html><body style='background:white'><p style='color:white;background:white;" +
            compositingStyle + "'>Composited visible text.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.LowContrastText
            && item.TextPreview.Contains("Composited visible text", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Throws<InvalidOperationException>(() => MhtmlDocument.RemoveSelectedContent(
            input,
            new OfficeContentCleanupSelection(new[] { finding.Id })));
    }

    [Fact]
    public void Mhtml_LinkedCompositedLowContrastIsReportOnly() {
        byte[] input = new MhtmlDocument(
            "<html><head><link rel='stylesheet' href='styles/site.css'></head>" +
            "<body><p class='target'>Linked composited text.</p></body></html>",
            new[] {
                new MhtmlResource(
                    Encoding.UTF8.GetBytes(".target{color:white;background:white;filter:invert(1)}"),
                    "text/css",
                    contentLocation: "styles/site.css")
            },
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.LowContrastText
            && item.TextPreview.Contains("Linked composited text", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Fact]
    public void Mhtml_NoOpCompositingKeepsLowContrastCleanable() {
        byte[] input = new MhtmlDocument(
            "<html><body style='background:white'><p style='color:white;background:white;" +
            "mix-blend-mode:normal;text-shadow:none;filter:none;-webkit-filter:none;backdrop-filter:none'>" +
            "No-op compositing text.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Kind == OfficeContentConcealmentKind.LowContrastText
            && item.TextPreview.Contains("No-op compositing text", StringComparison.Ordinal));

        Assert.NotEqual(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    [Theory]
    [InlineData("@scope")]
    [InlineData("@\\73 cope")]
    public void Mhtml_RejectsUnmodeledScopeRules(string scopeAtRule) {
        byte[] input = new MhtmlDocument(
            "<html><head><style>#secret{display:none}" + scopeAtRule +
            " (.chapter){#secret{display:block}}</style></head>" +
            "<body><div class='chapter'><p id='secret'>Scoped reveal.</p></div></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        Assert.Throws<InvalidDataException>(() => MhtmlDocument.InspectContentSafety(input));
    }

    [Fact]
    public void Mhtml_SelectableAlternateStylesheetConcealmentIsReportOnly() {
        byte[] input = new MhtmlDocument(
            "<html><head><link rel='stylesheet' title='default' href='default.css'>" +
            "<link rel='alternate stylesheet' title='readable' href='readable.css'></head>" +
            "<body><p class='switchable' title='Alternate generated title'>Alternate set reveal.</p></body></html>",
            new[] {
                new MhtmlResource(Encoding.UTF8.GetBytes(".switchable{display:none}"), "text/css", contentLocation: "default.css"),
                new MhtmlResource(Encoding.UTF8.GetBytes(
                    ".switchable{display:block}.switchable::before{content:attr(title)}"),
                    "text/css", contentLocation: "readable.css")
            },
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyReport report = MhtmlDocument.InspectContentSafety(input);
        OfficeContentSafetyFinding finding = Assert.Single(report.Findings, item =>
            item.TextPreview.Contains("Alternate set reveal", StringComparison.Ordinal));

        Assert.Equal(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
        Assert.Contains(report.Findings, item =>
            item.Location.EndsWith("/@title", StringComparison.Ordinal)
            && item.CleanupCapability == OfficeContentCleanupCapability.ReportOnly);
    }

    [Theory]
    [InlineData("attr(title)")]
    [InlineData("\\61 ttr(title)")]
    public void Mhtml_GeneratedContentAttributeIsReportOnly(string contentExpression) {
        byte[] input = new MhtmlDocument(
            "<html><head><style>p::before{content:" + contentExpression + "}</style></head>" +
            "<body><p title='Generated title'>Body.</p></body></html>",
            contentLocation: "https://example.test/index.html").ToBytes();

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.Location.EndsWith("/@title", StringComparison.Ordinal)
            && item.TextPreview.Contains("Generated title", StringComparison.Ordinal));

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

    [Fact]
    public void Mhtml_RejectsMalformedSelectedRootContentLocation() {
        byte[] input = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\n" +
            "Snapshot-Content-Location: https://example.test/fallback/index.html\r\n" +
            "Content-Type: multipart/related; boundary=outer\r\n\r\n" +
            "--outer\r\n" +
            "Content-Type: text/html; charset=utf-8\r\n" +
            "Content-Location: http://[invalid\r\n\r\n" +
            "<html><body><p style='display:none'>Malformed root location.</p></body></html>\r\n" +
            "--outer--\r\n");

        using (var stream = new MemoryStream(input, writable: false)) {
            MhtmlDocument document = MhtmlDocument.Load(stream);
            Assert.Contains(document.MimeDiagnostics, diagnostic =>
                diagnostic.Code == "MHTML_RESOURCE_CONTENT_LOCATION_INVALID" && diagnostic.Location == "root");
        }
        Assert.Throws<InvalidDataException>(() => MhtmlDocument.InspectContentSafety(input));
    }

    [Fact]
    public void Mhtml_ResolvesRelativeRootAgainstSnapshotLocation() {
        byte[] input = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\n" +
            "Snapshot-Content-Location: https://example.test/archive/\r\n" +
            "Content-Type: multipart/related; boundary=outer\r\n\r\n" +
            "--outer\r\nContent-Type: text/html; charset=utf-8\r\n" +
            "Content-Location: pages/index.html\r\n\r\n" +
            "<html><head><link rel='stylesheet' href='styles/site.css'></head>" +
            "<body><p class='target'>Snapshot-visible text.</p></body></html>\r\n" +
            "--outer\r\nContent-Type: text/css\r\n" +
            "Content-Location: https://example.test/archive/pages/styles/site.css\r\n\r\n" +
            ".target{display:block}\r\n" +
            "--outer\r\nContent-Type: text/css\r\n" +
            "Content-Location: mhtml://archive/pages/styles/site.css\r\n\r\n" +
            ".target{display:none}\r\n" +
            "--outer--\r\n");

        using (var stream = new MemoryStream(input, writable: false)) {
            MhtmlDocument document = MhtmlDocument.Load(stream);
            Assert.Equal("https://example.test/archive/pages/index.html", document.BaseUri.AbsoluteUri);
        }
        Assert.DoesNotContain(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Snapshot-visible text", StringComparison.Ordinal));
    }

    [Fact]
    public void Mhtml_RejectsInvalidRelatedRootContentIdentifier() {
        byte[] input = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/related; boundary=outer; start=\"<bad id>\"\r\n\r\n" +
            "--outer\r\nContent-Type: text/plain\r\n\r\nDefault root.\r\n" +
            "--outer\r\nContent-Type: text/html; charset=utf-8\r\nContent-ID: <bad id>\r\n\r\n" +
            "<html><body><p style='display:none'>Invalid identifier root.</p></body></html>\r\n" +
            "--outer--\r\n");

        using (EmailReadResult result = new EmailDocumentReader().Read(input)) {
            Assert.Contains(result.Diagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_MIME_CONTENT_ID_INVALID");
        }
        Assert.Throws<InvalidDataException>(() => MhtmlDocument.Load(new MemoryStream(input, writable: false)));
    }

    [Fact]
    public void Mhtml_IgnoresAmbiguityInsideOpaqueNestedMessageAttachment() {
        byte[] input = Encoding.ASCII.GetBytes(
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/related; boundary=outer\r\n\r\n" +
            "--outer\r\n" +
            "Content-Type: text/html; charset=utf-8\r\n\r\n" +
            "<html><body><p style='display:none'>Opaque nested diagnostics.</p></body></html>\r\n" +
            "--outer\r\n" +
            "Content-Type: message/rfc822; name=nested.eml\r\n" +
            "Content-Disposition: attachment; filename=nested.eml\r\n\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: multipart/mixed; boundary=inner\r\n" +
            "Content-Type: multipart/mixed; boundary=inner\r\n\r\n" +
            "--inner\r\nContent-Type: text/plain\r\n\r\nNested body.\r\n" +
            "--outer--\r\n");

        using (var stream = new MemoryStream(input, writable: false)) {
            MhtmlDocument document = MhtmlDocument.Load(stream);
            Assert.Contains(document.MimeDiagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_MIME_SINGLETON_HEADER_DUPLICATE"
                && diagnostic.Location?.Contains("/message", StringComparison.Ordinal) == true);
            Assert.Contains(document.MimeDiagnostics, diagnostic =>
                diagnostic.Code == "EMAIL_MIME_BOUNDARY_NOT_CLOSED"
                && diagnostic.Location?.Contains("/message", StringComparison.Ordinal) == true);
        }

        OfficeContentSafetyFinding finding = Assert.Single(MhtmlDocument.InspectContentSafety(input).Findings, item =>
            item.TextPreview.Contains("Opaque nested diagnostics", StringComparison.Ordinal));
        Assert.NotEqual(OfficeContentCleanupCapability.ReportOnly, finding.CleanupCapability);
    }

    private static byte[] BuildMhtmlWithBoundary(string declaredBoundary, string wireBoundary) => Encoding.UTF8.GetBytes(
        "MIME-Version: 1.0\r\n" +
        "Content-Type: multipart/related; boundary=\"" + declaredBoundary + "\"\r\n\r\n" +
        "--" + wireBoundary + "\r\n" +
        "Content-Type: text/html; charset=utf-8\r\n\r\n" +
        "<html><body><p style='display:none'>Invalid boundary.</p></body></html>\r\n" +
        "--" + wireBoundary + "--\r\n");
}
