#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
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
    public void ManagedProviderRejectsEncapsulatedContentForDetachedPdfSignature() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        PdfExternalSignaturePreparation preparation = CreateSigningPreparation();
        byte[] cms = CreateCustomCms(certificate, Encoding.UTF8.GetBytes("unrelated encapsulated content"), encapsulate: true, includeMessageDigest: true, useEcdsa: false);
        byte[] signedPdf = PdfIncrementalUpdater.ApplyExternalSignature(preparation, cms);
        var provider = new PdfPkcsSignatureCryptographyProvider(new PdfPkcsSignatureValidationOptions { ChainEvaluator = (_, _) => true });

        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(signedPdf, provider);
        PdfSignatureCryptographicResult result = Assert.Single(report.Signatures).CryptographicResult!;

        Assert.Equal(PdfCryptographicValidationStatus.Invalid, result.MathematicalSignatureStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Invalid, result.MessageDigestStatus);
        Assert.Contains(result.Findings, finding => finding.Code == "CmsDetachedContentExpected");
    }

    [Fact]
    public void ManagedProviderRejectsSignedAttributesWithoutMessageDigest() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        PdfExternalSignaturePreparation preparation = CreateSigningPreparation();
        byte[] cms = CreateCustomCms(certificate, preparation.SignedContent, encapsulate: false, includeMessageDigest: false, useEcdsa: false);
        byte[] signedPdf = PdfIncrementalUpdater.ApplyExternalSignature(preparation, cms);
        var provider = new PdfPkcsSignatureCryptographyProvider(new PdfPkcsSignatureValidationOptions { ChainEvaluator = (_, _) => true });

        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(signedPdf, provider);
        PdfSignatureCryptographicResult result = Assert.Single(report.Signatures).CryptographicResult!;

        Assert.Equal(PdfCryptographicValidationStatus.Invalid, result.MathematicalSignatureStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Invalid, result.MessageDigestStatus);
        Assert.Contains(result.Findings, finding => finding.Code == "CmsMessageDigestMissing");
    }

    [Fact]
    public void ManagedProviderValidatesDetachedEcdsaCmsSignature() {
        using X509Certificate2 certificate = CreateEcdsaSigningCertificate();
        PdfExternalSignaturePreparation preparation = CreateSigningPreparation();
        byte[] cms = CreateCustomCms(certificate, preparation.SignedContent, encapsulate: false, includeMessageDigest: true, useEcdsa: true);
        byte[] signedPdf = PdfIncrementalUpdater.ApplyExternalSignature(preparation, cms);
        var provider = new PdfPkcsSignatureCryptographyProvider(new PdfPkcsSignatureValidationOptions { ChainEvaluator = (_, _) => true });

        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(signedPdf, provider);
        PdfSignatureCryptographicResult result = Assert.Single(report.Signatures).CryptographicResult!;

        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.MathematicalSignatureStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.MessageDigestStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.CertificateChainStatus);
    }

    [Theory]
    [InlineData("20260714123456.7Z", 7000000L)]
    [InlineData("20260714123456.789Z", 7890000L)]
    [InlineData("20260714123456.789123456Z", 7891234L)]
    public void ManagedTimestampParserAcceptsFractionalGeneralizedTime(string encodedTime, long fractionalTicks) {
        byte[] encoded = PdfDerCodec.Wrap(0x18, Encoding.ASCII.GetBytes(encodedTime));

        DateTimeOffset? result = PdfManagedCmsDocument.ReadTime(new PdfDerReader(encoded).Read());

        Assert.Equal(new DateTimeOffset(2026, 7, 14, 12, 34, 56, TimeSpan.Zero).AddTicks(fractionalTicks), result);
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
        PdfExternalSignaturePreparation preparation = CreateSigningPreparation();
        using var signer = new PdfPkcsExternalSigner(certificate) { SigningTime = DateTimeOffset.UtcNow };
        return PdfIncrementalUpdater.ApplyExternalSignature(
            preparation,
            signer.Sign(new PdfExternalSignatureRequest(preparation)));
    }

    private static PdfExternalSignaturePreparation CreateSigningPreparation() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("CMS validation source"))
            .ToBytes();
        return PdfIncrementalUpdater.PrepareExternalSignature(
            source,
            new PdfExternalSignatureOptions {
                FieldName = "ApprovalSignature",
                Name = "OfficeIMO Test Signer",
                ReservedSignatureContentsBytes = 8192
            });
    }

    private static byte[] CreateCustomCms(
        X509Certificate2 certificate,
        byte[] content,
        bool encapsulate,
        bool includeMessageDigest,
        bool useEcdsa) {
        const string dataOid = "1.2.840.113549.1.7.1";
        const string signedDataOid = "1.2.840.113549.1.7.2";
        const string contentTypeAttributeOid = "1.2.840.113549.1.9.3";
        const string messageDigestAttributeOid = "1.2.840.113549.1.9.4";
        const string sha256Oid = "2.16.840.1.101.3.4.2.1";
        const string rsaEncryptionOid = "1.2.840.113549.1.1.1";
        const string ecdsaWithSha256Oid = "1.2.840.10045.4.3.2";

        byte[] digest = SHA256.HashData(content);
        var attributes = new List<byte[]> {
            PdfDerCodec.Sequence(PdfDerCodec.ObjectIdentifier(contentTypeAttributeOid), PdfDerCodec.Set(PdfDerCodec.ObjectIdentifier(dataOid)))
        };
        if (includeMessageDigest) {
            attributes.Add(PdfDerCodec.Sequence(PdfDerCodec.ObjectIdentifier(messageDigestAttributeOid), PdfDerCodec.Set(PdfDerCodec.OctetString(digest))));
        }
        byte[] signedAttributes = PdfDerCodec.Set(attributes.ToArray());
        byte[] signature;
        string signatureAlgorithmOid;
        if (useEcdsa) {
            using ECDsa ecdsa = certificate.GetECDsaPrivateKey() ?? throw new InvalidOperationException("ECDSA test certificate has no private key.");
            signature = ecdsa.SignData(signedAttributes, HashAlgorithmName.SHA256, DSASignatureFormat.Rfc3279DerSequence);
            signatureAlgorithmOid = ecdsaWithSha256Oid;
        } else {
            using RSA rsa = certificate.GetRSAPrivateKey() ?? throw new InvalidOperationException("RSA test certificate has no private key.");
            signature = rsa.SignData(signedAttributes, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            signatureAlgorithmOid = rsaEncryptionOid;
        }

        byte[] serial = certificate.GetSerialNumber();
        Array.Reverse(serial);
        byte[] signerInfo = PdfDerCodec.Sequence(
            PdfDerCodec.Integer(1),
            PdfDerCodec.Sequence(certificate.IssuerName.RawData, PdfDerCodec.Integer(serial)),
            PdfDerCodec.AlgorithmIdentifier(sha256Oid),
            PdfDerCodec.ReplaceTag(signedAttributes, 0xA0),
            PdfDerCodec.AlgorithmIdentifier(signatureAlgorithmOid, includeNull: !useEcdsa),
            PdfDerCodec.OctetString(signature));
        byte[] contentInfo = encapsulate
            ? PdfDerCodec.Sequence(PdfDerCodec.ObjectIdentifier(dataOid), PdfDerCodec.Context(0, PdfDerCodec.OctetString(content)))
            : PdfDerCodec.Sequence(PdfDerCodec.ObjectIdentifier(dataOid));
        byte[] signedData = PdfDerCodec.Sequence(
            PdfDerCodec.Integer(1),
            PdfDerCodec.Set(PdfDerCodec.AlgorithmIdentifier(sha256Oid)),
            contentInfo,
            PdfDerCodec.Context(0, certificate.RawData),
            PdfDerCodec.Set(signerInfo));
        return PdfDerCodec.Sequence(PdfDerCodec.ObjectIdentifier(signedDataOid), PdfDerCodec.Context(0, signedData));
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

    private static X509Certificate2 CreateEcdsaSigningCertificate() {
        using ECDsa ecdsa = ECDsa.Create(ECCurve.NamedCurves.nistP256);
        var request = new CertificateRequest("CN=OfficeIMO PDF ECDSA Test", ecdsa, HashAlgorithmName.SHA256);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }

}
#endif
