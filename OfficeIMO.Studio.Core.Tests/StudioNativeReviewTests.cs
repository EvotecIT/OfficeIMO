using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
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
                document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(300)).Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(text)))));
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
    public async Task Shapes_text_boxes_and_arrows_are_drawn_as_annotations() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("shapes.pdf", "Shapes"), null, default);
        var red = new StudioNativeColor(0.85, 0.12, 0.1);

        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("rectangle", 1, new StudioNativeRectangle(40, 60, 100, 50), Color: red, StrokeWidth: 2), default);
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("ellipse", 1, new StudioNativeRectangle(160, 60, 80, 50), Color: red, StrokeWidth: 2), default);
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("freeText", 1, new StudioNativeRectangle(40, 140, 160, 30), Text: "Check this", Color: red), default);
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("line", 1, Strokes: [[new StudioNativePoint(260, 200), new StudioNativePoint(340, 150)]],
            Color: red, StrokeWidth: 2, LineEnd: "OpenArrow"), default);

        Assert.Equal(["Circle", "FreeText", "Line", "Square"],
            (await document.GetAnnotationsAsync(1, default)).Annotations.Select(annotation => annotation.Subtype).Order());
        string saved = Path.Combine(_root, "shapes-saved.pdf");
        await document.SaveCopyAsync(saved, default);
        IReadOnlyList<PdfAnnotation> annotations = PdfDocument.Load(saved).Annotations.GetForEditing();
        Assert.Equal("OpenArrow", Assert.Single(annotations, annotation => annotation.Subtype == "Line").LineEndEnding);
        Assert.Equal("Check this", Assert.Single(annotations, annotation => annotation.Subtype == "FreeText").Contents);
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.AddAnnotationAsync(new StudioNativeAnnotationRequest("rectangle", 1,
            new StudioNativeRectangle(40, 60, 100, 50), LineEnd: "OpenArrow"), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.AddAnnotationAsync(new StudioNativeAnnotationRequest("line", 1,
            Strokes: [[new StudioNativePoint(10, 10), new StudioNativePoint(50, 50)]], LineEnd: "Butt"), default));
    }

    [Fact]
    public async Task Annotations_move_and_resize_on_the_displayed_page_and_outlines_become_polygons() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("move.pdf", "Shapes"), null, default);
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("rectangle", 1, new StudioNativeRectangle(40, 60, 100, 50), StrokeWidth: 2), default);
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("polygon", 1, Strokes: [[new StudioNativePoint(200, 60),
            new StudioNativePoint(260, 70), new StudioNativePoint(250, 130), new StudioNativePoint(190, 120)]], StrokeWidth: 2), default);
        IReadOnlyList<StudioNativeAnnotation> added = (await document.GetAnnotationsAsync(1, default)).Annotations;
        StudioNativeAnnotation square = Assert.Single(added, annotation => annotation.Subtype == "Square");
        Assert.Contains(added, annotation => annotation.Subtype == "Polygon");

        StudioNativeRectangle moved = square.Bounds with { X = square.Bounds.X + 30, Y = square.Bounds.Y + 100 };
        await document.UpdateAnnotationAsync(new StudioNativeAnnotationUpdate(square.ObjectNumber, Bounds: moved), default);
        StudioNativeRectangle after = Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations, annotation => annotation.Subtype == "Square").Bounds;
        Assert.Equal(moved.X, after.X, 1);
        Assert.Equal(moved.Y, after.Y, 1);
        Assert.Equal(square.Bounds.Width, after.Width, 1);

        StudioNativeRectangle resized = after with { Width = after.Width * 2, Height = after.Height / 2 };
        await document.UpdateAnnotationAsync(new StudioNativeAnnotationUpdate(square.ObjectNumber, Bounds: resized), default);
        StudioNativeRectangle larger = Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations, annotation => annotation.Subtype == "Square").Bounds;
        Assert.Equal(resized.Width, larger.Width, 1);
        Assert.Equal(resized.Height, larger.Height, 1);

        await document.UndoAsync(default);
        await document.UndoAsync(default);
        Assert.Equal(square.Bounds, Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations, annotation => annotation.Subtype == "Square").Bounds);
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.UpdateAnnotationAsync(
            new StudioNativeAnnotationUpdate(square.ObjectNumber, Contents: "Both", Bounds: moved), default));
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
    public async Task Signature_heading_uses_the_certificate_subject_instead_of_the_Pdf_name() {
        string source = CreateTextPdf("spoof-source.pdf", "Contract");
        using RSA key = RSA.Create(2048);
        var request = new CertificateRequest("CN=Actual Signer", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.Now.AddDays(-1), DateTimeOffset.Now.AddDays(1));
        using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
        byte[] signed = PdfDocument.Load(source).Security.SignExternal(signer,
            new PdfExternalSignatureOptions { Name = "Impersonated Name", ReservedSignatureContentsBytes = 8192 }).Pdf;
        string path = Path.Combine(_root, "spoofed-name.pdf");
        File.WriteAllBytes(path, signed);

        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);
        StudioNativeSignature signature = Assert.Single((await document.CheckAsync(new StudioNativeCheckRequest(Signatures: true), default)).Signatures!.Signatures);
        Assert.Equal("Actual Signer", signature.Signer);
        Assert.NotEqual("Impersonated Name", signature.Signer);
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
    public async Task Health_copies_reject_a_source_alias_without_changing_source_bytes() {
        string source = CreateTextPdf("health-alias.pdf", "One");
        string alias = Path.Combine(_root, "source-link.pdf");
        File.CreateSymbolicLink(alias, source);
        byte[] original = File.ReadAllBytes(source);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(source, null, default);
        await Assert.ThrowsAsync<ArgumentException>(() => document.WriteAsync(
            new StudioNativeWriteRequest(Optimize: new StudioNativeHealthCopy(alias)), default));
        Assert.Equal(original, File.ReadAllBytes(source));
        Assert.True(new FileInfo(alias).LinkTarget is not null);
    }

    [Fact]
    public async Task Encrypted_owner_rectangle_redaction_retains_authentication_and_undo() {
        string path = CreateTextPdf("encrypted-redact.pdf", "Account 4111 1111 belongs to Alex");
        File.WriteAllBytes(path, PdfDocument.Load(path).Security.Encrypt(new("reader") {
            OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.None
        }).Pdf);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, "owner", default);
        Assert.True((await document.GetStructureAsync(default)).CanRedact);
        StudioNativeRectangle name = Assert.Single((await document.RedactAsync(
            new(Find: new("Alex")), default)).Marks!).Bounds!;
        StudioNativeRedactionVerification result = (await document.RedactAsync(new(Apply: new(
            [new(1, name with { X = name.X - 2, Y = name.Y - 2, Width = name.Width + 4, Height = name.Height + 4 })],
            Revision: (await document.GetInfoAsync(default)).Revision)), default)).Verification!;
        Assert.True(result.Verified);
        Assert.DoesNotContain("Alex", await PageTextAsync(document, 1));
        Assert.Contains("Account", await PageTextAsync(document, 1));
        Assert.True((await document.GetInfoAsync(default)).IsEncrypted);
        await document.UndoAsync(default);
        Assert.Contains("Alex", await PageTextAsync(document, 1));
    }

    [Fact]
    public async Task Redaction_finds_text_removes_it_verifiably_and_undoes_as_one_change() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("redact.pdf", "Account 4111 1111 belongs to Alex", "Nothing here"), null, default);

        IReadOnlyList<StudioNativeRedactionMark> marks = (await document.RedactAsync(
            new StudioNativeRedactionRequest(Find: new StudioNativeRedactionFind("4111 1111")), default)).Marks!;
        StudioNativeRedactionMark mark = Assert.Single(marks);
        Assert.Equal(1, mark.Page);

        StudioNativeRedactionVerification verification = (await document.RedactAsync(new StudioNativeRedactionRequest(
            Apply: new StudioNativeRedactionApply(Find: new StudioNativeRedactionFind("4111 1111"), Include: [0], RemovedText: "4111 1111",
                Revision: (await document.GetInfoAsync(default)).Revision)), default)).Verification!;

        Assert.True(verification.Verified);
        Assert.Equal([1], verification.Pages);
        string text = await PageTextAsync(document, 1);
        Assert.DoesNotContain("4111", text);
        Assert.Contains("Alex", text);
        await document.UndoAsync(default);
        Assert.Contains("4111", await PageTextAsync(document, 1));

        // An area drawn over the page removes what it covers, as a drag across the page does.
        StudioNativeRectangle name = Assert.Single((await document.RedactAsync(
            new StudioNativeRedactionRequest(Find: new StudioNativeRedactionFind("Alex")), default)).Marks!).Bounds!;
        StudioNativeRedactionVerification drawn = (await document.RedactAsync(new StudioNativeRedactionRequest(Apply: new StudioNativeRedactionApply(
            [new StudioNativeRedactionMark(1, name with { X = name.X - 2, Y = name.Y - 2, Width = name.Width + 4, Height = name.Height + 4 })],
            Revision: (await document.GetInfoAsync(default)).Revision)), default)).Verification!;
        Assert.True(drawn.Verified);
        text = await PageTextAsync(document, 1);
        Assert.DoesNotContain("Alex", text);
        Assert.Contains("Account", text);
        await document.UndoAsync(default);

        long currentRevision = (await document.GetInfoAsync(default)).Revision;
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.RedactAsync(new StudioNativeRedactionRequest(
            Apply: new StudioNativeRedactionApply([new StudioNativeRedactionMark(1, new StudioNativeRectangle(10, 10, 0, 5))],
                Revision: currentRevision)), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.RedactAsync(new StudioNativeRedactionRequest(
            Apply: new StudioNativeRedactionApply(Find: new StudioNativeRedactionFind("4111 1111"), Include: [3],
                Revision: currentRevision)), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.RedactAsync(new StudioNativeRedactionRequest(), default));
    }
}
