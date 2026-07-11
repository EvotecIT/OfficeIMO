#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Cryptography;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfExternalEngineProofTests {
    [Fact]
    public void GenerateSyntaxRenderingAndSignatureProofFixtures() {
        byte[] source = PdfDocument.Create()
            .Meta(title: "OfficeIMO external engine proof")
            .Paragraph(paragraph => paragraph.Text("External syntax and rendering proof"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second proof page"))
            .ToBytes();
        byte[] rewritten = PdfPageEditor.ReversePages(
            PdfMetadataEditor.UpdateMetadata(source, subject: "Mutation proof"));

        PdfRewritePreservationReport preservation = PdfRewritePreservation.Assess(
            source,
            rewritten,
            new PdfRewritePreservationOptions {
                PreserveRevisionStructure = false
            }.AllowMetadataChanges("Subject"));
        Assert.True(preservation.IsPreserved, preservation.Summary);
        Assert.Equal(2, PdfInspector.Inspect(rewritten).PageCount);

        using X509Certificate2 certificate = CreateSigningCertificate();
        byte[] signed = Sign(rewritten, certificate);
        var provider = new PdfPkcsSignatureCryptographyProvider(new PdfPkcsSignatureValidationOptions {
            ChainEvaluator = (_, _) => true
        });
        PdfSignatureValidationReport signatureReport = PdfSignatureValidator.Validate(signed, provider);
        Assert.True(signatureReport.IsStructurallyValid);
        Assert.True(signatureReport.MathematicalSignaturesVerified);
        Assert.True(signatureReport.DigestVerified);

        string? outputDirectory = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_ENGINE_PROOF_OUTPUT");
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            return;
        }

        Directory.CreateDirectory(outputDirectory);
        File.WriteAllBytes(Path.Combine(outputDirectory, "officeimo-engine-rewrite.pdf"), rewritten);
        File.WriteAllBytes(Path.Combine(outputDirectory, "officeimo-engine-signed.pdf"), signed);
    }

    private static byte[] Sign(byte[] source, X509Certificate2 certificate) {
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Approval,
                FieldName = "EngineProofSignature",
                Name = "OfficeIMO Engine Proof",
                Reason = "External signature validation proof",
                VisibleAppearance = new PdfVisibleSignatureAppearanceOptions {
                    PageNumber = 1,
                    Text = "OfficeIMO Engine Proof"
                },
                ReservedSignatureContentsBytes = 8192
            });
        var cms = new SignedCms(new ContentInfo(preparation.SignedContent), detached: true);
        var signer = new CmsSigner(SubjectIdentifierType.IssuerAndSerialNumber, certificate) {
            IncludeOption = X509IncludeOption.EndCertOnly
        };
        cms.ComputeSignature(signer);
        return PdfIncrementalUpdater.ApplyExternalSignature(preparation, cms.Encode());
    }

    private static X509Certificate2 CreateSigningCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=OfficeIMO PDF Engine Proof",
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }
}
#endif
