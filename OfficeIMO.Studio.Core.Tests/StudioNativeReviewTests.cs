using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

/// <summary>Links, stamps, checks, health copies and verified redaction through the engine library the Apple apps call.</summary>
public sealed class StudioNativeReviewTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-native-review-" + Guid.NewGuid().ToString("N"));

    public StudioNativeReviewTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }

    public void Dispose() => Directory.Delete(_root, recursive: true);

    private string CreateTextPdf(string name, params string[] pages) {
        string path = Path.Combine(_root, name);
        PdfDocument.Create(document => {
            foreach (string text in pages)
                document.Page(page => page.Size(400, 300).Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(text)))));
        }).Save(path);
        return path;
    }

    private static async Task<string> PageTextAsync(StudioNativeDocument document, int page) =>
        string.Concat((await document.GetPageTextAsync(page, default)).Texts);

    [Fact]
    public async Task Links_go_to_a_web_address_or_a_page_and_stamps_use_standard_names() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("links.pdf", "One", "Two"), null, default);
        var area = new StudioNativeRectangle(40, 40, 120, 20);

        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("link", 1, area, LinkUri: "https://officeimo.com/"), default);
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("link", 1, area with { Y = 80 }, LinkPage: 2), default);
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("stamp", 2, new StudioNativeRectangle(100, 100, 150, 50), Stamp: "Approved"), default);

        Assert.Equal(2, (await document.GetAnnotationsAsync(1, default)).Annotations.Count(annotation => annotation.Subtype == "Link"));
        Assert.Contains((await document.GetAnnotationsAsync(2, default)).Annotations, annotation => annotation.Subtype == "Stamp");
        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.AddAnnotationAsync(new StudioNativeAnnotationRequest("link", 1, area, LinkUri: "javascript:alert(1)"), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.AddAnnotationAsync(new StudioNativeAnnotationRequest("link", 1, area, LinkUri: "https://officeimo.com/", LinkPage: 2), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.AddAnnotationAsync(new StudioNativeAnnotationRequest("link", 1, area, LinkPage: 9), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.AddAnnotationAsync(new StudioNativeAnnotationRequest("stamp", 1, area, Stamp: "Whatever"), default));
    }

    [Fact]
    public async Task Signatures_are_reported_for_a_signed_copy_and_absent_from_an_unsigned_document() {
        string source = CreateTextPdf("unsigned.pdf", "Contract");
        using (StudioNativeDocument unsigned = await StudioNativeDocument.OpenAsync(source, null, default)) {
            StudioNativeSignatureReport none = (await unsigned.CheckAsync(new StudioNativeCheckRequest(Signatures: true), default)).Signatures!;
            Assert.False(none.HasSignatures);
            Assert.Empty(none.Signatures);

            using RSA key = RSA.Create(2048);
            var request = new CertificateRequest("CN=Studio Signer", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.NonRepudiation, true));
            using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.Now.AddDays(-1), DateTimeOffset.Now.AddYears(1));
            string identity = Path.Combine(_root, "signer.p12");
            File.WriteAllBytes(identity, certificate.Export(X509ContentType.Pkcs12, "secret"));
            await unsigned.SaveSignedCopyAsync(new StudioNativeSigningRequest(Path.Combine(_root, "signed.pdf"), identity, "secret", "Approved", "Kraków"), default);
        }

        using StudioNativeDocument signed = await StudioNativeDocument.OpenAsync(Path.Combine(_root, "signed.pdf"), null, default);
        StudioNativeSignatureReport report = (await signed.CheckAsync(new StudioNativeCheckRequest(Signatures: true), default)).Signatures!;
        StudioNativeSignature signature = Assert.Single(report.Signatures);
        Assert.True(report.HasSignatures);
        Assert.Contains(signature.Status, new[] { "valid", "intact" });
        Assert.True(signature.CoversWholeDocument);
        Assert.Equal(("Studio Signer", "Approved", "Kraków"), (signature.Signer, signature.Reason, signature.Location));
    }

    [Fact]
    public async Task Standards_and_health_are_reported_without_changing_the_document() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("check.pdf", "One", "Two"), null, default);

        StudioNativeStandardReport standard = (await document.CheckAsync(new StudioNativeCheckRequest(Standard: "pdfua1"), default)).Standard!;
        Assert.NotEmpty(standard.Requirements);
        Assert.All(standard.Requirements, requirement => Assert.Contains(requirement.Status, new[] { "met", "missing", "unsupported" }));
        Assert.Equal(standard.Requirements.All(requirement => requirement.Status == "met"), standard.IsReady);

        StudioNativeHealth health = (await document.CheckAsync(new StudioNativeCheckRequest(Health: true), default)).Health!;
        Assert.Equal(2, health.Pages);
        Assert.True(health.CanRead);
        Assert.False(health.IsEncrypted || health.IsSigned);

        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.CheckAsync(new StudioNativeCheckRequest(Standard: "iso9001"), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.CheckAsync(new StudioNativeCheckRequest(Signatures: true, Health: true), default));
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
    }

    [Fact]
    public async Task Optimized_and_sanitized_copies_are_written_beside_the_document() {
        string source = CreateTextPdf("health.pdf", "One", "Two");
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(source, null, default);
        string optimized = Path.Combine(_root, "optimized.pdf");
        string sanitized = Path.Combine(_root, "sanitized.pdf");

        StudioNativeWriteResult optimization = await document.WriteAsync(new StudioNativeWriteRequest(Optimize: new StudioNativeHealthCopy(optimized, "smallest")), default);
        StudioNativeWriteResult sanitization = await document.WriteAsync(new StudioNativeWriteRequest(Sanitize: new StudioNativeHealthCopy(sanitized)), default);

        Assert.Equal([optimized], optimization.Files);
        Assert.False(string.IsNullOrEmpty(optimization.Summary));
        Assert.False(string.IsNullOrEmpty(sanitization.Summary));
        foreach (string copy in new[] { optimized, sanitized }) {
            using StudioNativeDocument reopened = await StudioNativeDocument.OpenAsync(copy, null, default);
            Assert.Equal(2, (await reopened.GetInfoAsync(default)).Pages.Count);
            Assert.Contains("Two", await PageTextAsync(reopened, 2));
        }
        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.WriteAsync(new StudioNativeWriteRequest(Optimize: new StudioNativeHealthCopy(source)), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.WriteAsync(new StudioNativeWriteRequest(Optimize: new StudioNativeHealthCopy(optimized, "tiny")), default));
    }

    [Fact]
    public async Task Redaction_finds_text_removes_it_verifiably_and_undoes_as_one_change() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("redact.pdf", "Account 4111 1111 belongs to Alex", "Nothing here"), null, default);

        IReadOnlyList<StudioNativeRedactionMark> marks = (await document.RedactAsync(
            new StudioNativeRedactionRequest(Find: new StudioNativeRedactionFind("4111 1111")), default)).Marks!;
        StudioNativeRedactionMark mark = Assert.Single(marks);
        Assert.Equal(1, mark.Page);

        StudioNativeRedactionVerification verification = (await document.RedactAsync(new StudioNativeRedactionRequest(
            Apply: new StudioNativeRedactionApply(Find: new StudioNativeRedactionFind("4111 1111"), Include: [0], RemovedText: "4111 1111")), default)).Verification!;

        Assert.True(verification.Verified);
        Assert.Equal([1], verification.Pages);
        string text = await PageTextAsync(document, 1);
        Assert.DoesNotContain("4111", text);
        Assert.Contains("Alex", text);
        await document.UndoAsync(default);
        Assert.Contains("4111", await PageTextAsync(document, 1));

        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.RedactAsync(new StudioNativeRedactionRequest(
            Apply: new StudioNativeRedactionApply([new StudioNativeRedactionMark(1, new StudioNativeRectangle(10, 10, 0, 5))])), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.RedactAsync(new StudioNativeRedactionRequest(
            Apply: new StudioNativeRedactionApply(Find: new StudioNativeRedactionFind("4111 1111"), Include: [3])), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.RedactAsync(new StudioNativeRedactionRequest(), default));
    }
}
