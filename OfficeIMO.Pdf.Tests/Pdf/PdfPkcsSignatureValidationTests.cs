#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Cryptography;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPkcsSignatureValidationTests {
    [Fact]
    public void SignedCmsProviderValidatesDetachedPdfSignatureAndCallerTrustPolicy() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        byte[] signedPdf = CreateSignedPdf(certificate);
        var provider = new PdfPkcsSignatureCryptographyProvider(new PdfPkcsSignatureValidationOptions {
            ChainEvaluator = (_, _) => true
        });

        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(signedPdf, provider);
        PdfSignatureValidationResult signature = Assert.Single(report.Signatures);
        PdfSignatureCryptographicResult cryptographic = Assert.IsType<PdfSignatureCryptographicResult>(signature.CryptographicResult);

        Assert.True(report.IsStructurallyValid);
        Assert.True(report.CryptographicValidationPerformed);
        Assert.True(report.MathematicalSignaturesVerified);
        Assert.True(report.DigestVerified);
        Assert.True(report.CertificateChainVerified);
        Assert.True(report.CryptographicTrustVerified);
        Assert.False(report.RevocationChecked);
        Assert.False(report.TimestampValidationPerformed);
        Assert.Equal("CryptographicallyValidAndTrusted", report.ProofStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Valid, cryptographic.MathematicalSignatureStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Valid, cryptographic.MessageDigestStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Valid, cryptographic.CertificateChainStatus);
        Assert.Equal(PdfCryptographicValidationStatus.NotPerformed, cryptographic.RevocationStatus);
        Assert.Equal(certificate.Thumbprint, cryptographic.SignerThumbprint);
    }

    [Fact]
    public void SignedCmsProviderDetectsTamperedSignedByteRanges() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        byte[] signedPdf = CreateSignedPdf(certificate);
        signedPdf[7] = signedPdf[7] == (byte)'4' ? (byte)'5' : (byte)'4';
        var provider = new PdfPkcsSignatureCryptographyProvider(new PdfPkcsSignatureValidationOptions {
            ChainEvaluator = (_, _) => true
        });

        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(signedPdf, provider);
        PdfSignatureCryptographicResult cryptographic = Assert.Single(report.Signatures).CryptographicResult!;

        Assert.False(report.MathematicalSignaturesVerified);
        Assert.False(report.DigestVerified);
        Assert.Equal("CryptographicInvalid", report.ProofStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Invalid, cryptographic.MathematicalSignatureStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Invalid, cryptographic.MessageDigestStatus);
        Assert.Contains(report.Findings, finding => finding.Code == "CmsSignatureInvalid");
    }

    [Fact]
    public void ExternalSignerContractCompletesAndValidatesApprovalSignature() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("External signer callback source"))
            .ToBytes();
        using var signer = new PdfPkcsExternalSigner(certificate, "Test local certificate signer");

        PdfExternalSignatureCompletion completion = PdfIncrementalUpdater.SignExternal(
            source,
            signer,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Approval,
                FieldName = "CloudApproval",
                ReservedSignatureContentsBytes = 8192
            });
        var provider = new PdfPkcsSignatureCryptographyProvider(new PdfPkcsSignatureValidationOptions {
            ChainEvaluator = (_, _) => true
        });
        PdfSignatureValidationReport report = completion.ToDocument().ValidateSignatures(provider);

        Assert.Equal(signer.Name, completion.SignerName);
        Assert.True(completion.SignatureContentsLength > 0);
        Assert.Equal(PdfSignatureProfile.Approval, completion.Preparation.Profile);
        Assert.True(report.MathematicalSignaturesVerified);
        Assert.True(report.DigestVerified);
    }

    private static byte[] CreateSignedPdf(X509Certificate2 certificate) {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("CMS validation source"))
            .ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                FieldName = "ApprovalSignature",
                Name = "OfficeIMO Test Signer",
                ReservedSignatureContentsBytes = 8192
            });
        using var signer = new PdfPkcsExternalSigner(certificate) { SigningTime = DateTimeOffset.UtcNow };
        return PdfIncrementalUpdater.ApplyExternalSignature(
            preparation,
            signer.Sign(new PdfExternalSignatureRequest(preparation)));
    }

    private static X509Certificate2 CreateSigningCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=OfficeIMO PDF Signature Test",
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }

}
#endif
