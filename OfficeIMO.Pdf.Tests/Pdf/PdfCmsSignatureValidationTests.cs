#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfCmsSignatureValidationTests {
    [Fact]
    public void SignedCmsProviderValidatesDetachedPdfSignatureAndCallerTrustPolicy() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        byte[] signedPdf = CreateSignedPdf(certificate);
        PdfCmsSignatureCryptographyProvider provider = CreateProvider((_, _) => true);

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
        PdfCmsSignatureCryptographyProvider provider = CreateProvider((_, _) => true);

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
        PdfCmsSignatureCryptographyProvider provider = CreateProvider((_, _) => true);

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
        PdfCmsSignatureCryptographyProvider provider = CreateProvider((_, _) => true);

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
        PdfCmsSignatureCryptographyProvider provider = CreateProvider((_, _) => true);

        PdfSignatureValidationReport report = PdfSignatureValidator.Validate(signedPdf, provider);
        PdfSignatureCryptographicResult result = Assert.Single(report.Signatures).CryptographicResult!;

        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.MathematicalSignatureStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.MessageDigestStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.CertificateChainStatus);
    }

    [Fact]
    public void ManagedProviderRejectsSignerIdentifierThatDoesNotMatchEmbeddedCertificate() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        using X509Certificate2 unrelatedIdentifier = CreateSigningCertificate();
        PdfExternalSignaturePreparation preparation = CreateSigningPreparation();
        byte[] cms = CreateCustomCms(certificate, preparation.SignedContent, encapsulate: false, includeMessageDigest: true, useEcdsa: false, signerIdentifierCertificate: unrelatedIdentifier);
        byte[] signedPdf = PdfIncrementalUpdater.ApplyExternalSignature(preparation, cms);
        PdfCmsSignatureCryptographyProvider provider = CreateProvider((_, _) => true);

        PdfSignatureCryptographicResult result = Assert.Single(PdfSignatureValidator.Validate(signedPdf, provider).Signatures).CryptographicResult!;

        Assert.Equal(PdfCryptographicValidationStatus.Invalid, result.MathematicalSignatureStatus);
        Assert.Contains(result.Findings, finding => finding.Code == "CmsSignerMissing");
    }

    [Fact]
    public void ManagedProviderMatchesSignerBySubjectKeyIdentifier() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        PdfExternalSignaturePreparation preparation = CreateSigningPreparation();
        byte[] cms = CreateCustomCms(certificate, preparation.SignedContent, encapsulate: false, includeMessageDigest: true, useEcdsa: false, useSubjectKeyIdentifier: true);
        byte[] signedPdf = PdfIncrementalUpdater.ApplyExternalSignature(preparation, cms);
        PdfCmsSignatureCryptographyProvider provider = CreateProvider((_, _) => true);

        PdfSignatureCryptographicResult result = Assert.Single(PdfSignatureValidator.Validate(signedPdf, provider).Signatures).CryptographicResult!;

        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.MathematicalSignatureStatus);
        Assert.Equal(certificate.Thumbprint, result.SignerThumbprint);
    }

    [Fact]
    public void SignatureTimestampRejectsUntrustedTsaCertificateChain() {
        using X509Certificate2 signerCertificate = CreateSigningCertificate();
        using X509Certificate2 tsaCertificate = CreateTimestampCertificate();
        PdfExternalSignaturePreparation preparation = CreateSigningPreparation();
        byte[] cms = CreateCustomCms(
            signerCertificate,
            preparation.SignedContent,
            encapsulate: false,
            includeMessageDigest: true,
            useEcdsa: false,
            signatureTimestampFactory: signature => CreateTimestampToken(tsaCertificate, signature));
        byte[] signedPdf = PdfIncrementalUpdater.ApplyExternalSignature(preparation, cms);
        PdfCmsSignatureCryptographyProvider provider = CreateProvider(
            (certificate, _) => !string.Equals(
                certificate.Thumbprint,
                tsaCertificate.Thumbprint,
                StringComparison.OrdinalIgnoreCase));

        PdfSignatureCryptographicResult result = Assert.Single(PdfSignatureValidator.Validate(signedPdf, provider).Signatures).CryptographicResult!;

        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.MathematicalSignatureStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Valid, result.CertificateChainStatus);
        Assert.Equal(PdfCryptographicValidationStatus.Invalid, result.TimestampStatus);
        Assert.Contains(result.Findings, finding => finding.Code == "CertificateChainUntrusted" && finding.Message.Contains("TSA", StringComparison.Ordinal));
        Assert.Contains(result.Findings, finding => finding.Code == "SignatureTimestampInvalid");
    }

    [Fact]
    public void ExternalSignerContractCompletesAndValidatesApprovalSignature() {
        using X509Certificate2 certificate = CreateSigningCertificate();
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("External signer callback source"))
            .ToBytes();
        using var signer = new PdfCmsExternalSigner(certificate, "Test local certificate signer");

        PdfExternalSignatureCompletion completion = PdfIncrementalUpdater.SignExternal(
            source,
            signer,
            new PdfExternalSignatureOptions {
                Profile = PdfSignatureProfile.Approval,
                FieldName = "CloudApproval",
                ReservedSignatureContentsBytes = 8192
            });
        PdfCmsSignatureCryptographyProvider provider = CreateProvider((_, _) => true);
        PdfSignatureValidationReport report = completion.ToDocument().ValidateSignatures(provider);

        Assert.Equal(signer.Name, completion.SignerName);
        Assert.True(completion.SignatureContentsLength > 0);
        Assert.Equal(PdfSignatureProfile.Approval, completion.Preparation.Profile);
        Assert.True(report.MathematicalSignaturesVerified);
        Assert.True(report.DigestVerified);
    }

    private static byte[] CreateSignedPdf(X509Certificate2 certificate) {
        PdfExternalSignaturePreparation preparation = CreateSigningPreparation();
        using var signer = new PdfCmsExternalSigner(
            certificate,
            signingOptions: new CmsSigningOptions { SigningTime = DateTimeOffset.UtcNow });
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
        bool useEcdsa,
        X509Certificate2? signerIdentifierCertificate = null,
        bool useSubjectKeyIdentifier = false,
        Func<byte[], byte[]>? signatureTimestampFactory = null) {
        const string dataOid = "1.2.840.113549.1.7.1";
        const string signedDataOid = "1.2.840.113549.1.7.2";
        const string contentTypeAttributeOid = "1.2.840.113549.1.9.3";
        const string messageDigestAttributeOid = "1.2.840.113549.1.9.4";
        const string signatureTimestampAttributeOid = "1.2.840.113549.1.9.16.2.14";
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

        X509Certificate2 signerIdentifier = signerIdentifierCertificate ?? certificate;
        byte[] serial = signerIdentifier.GetSerialNumber();
        Array.Reverse(serial);
        byte[] encodedSignerIdentifier = useSubjectKeyIdentifier
            ? PdfDerCodec.Wrap(
                0x80,
                Convert.FromHexString(Assert.IsType<X509SubjectKeyIdentifierExtension>(
                    Assert.Single(signerIdentifier.Extensions.Cast<X509Extension>(), extension => extension.Oid?.Value == "2.5.29.14")).SubjectKeyIdentifier!))
            : PdfDerCodec.Sequence(signerIdentifier.IssuerName.RawData, PdfDerCodec.Integer(serial));
        var signerInfoValues = new List<byte[]> {
            PdfDerCodec.Integer(useSubjectKeyIdentifier ? 3 : 1),
            encodedSignerIdentifier,
            PdfDerCodec.AlgorithmIdentifier(sha256Oid),
            PdfDerCodec.ReplaceTag(signedAttributes, 0xA0),
            PdfDerCodec.AlgorithmIdentifier(signatureAlgorithmOid, includeNull: !useEcdsa),
            PdfDerCodec.OctetString(signature)
        };
        if (signatureTimestampFactory != null) {
            byte[] timestampToken = signatureTimestampFactory(signature);
            byte[] timestampAttribute = PdfDerCodec.Sequence(
                PdfDerCodec.ObjectIdentifier(signatureTimestampAttributeOid),
                PdfDerCodec.Set(timestampToken));
            signerInfoValues.Add(PdfDerCodec.Context(1, timestampAttribute));
        }
        byte[] signerInfo = PdfDerCodec.Sequence(signerInfoValues.ToArray());
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

    private static byte[] CreateTimestampToken(X509Certificate2 certificate, byte[] expectedData) {
        using RSA rsa = certificate.GetRSAPrivateKey() ?? throw new InvalidOperationException("TSA test certificate has no private key.");
        Org.BouncyCastle.X509.X509Certificate bcCertificate =
            Org.BouncyCastle.Security.DotNetUtilities.FromX509Certificate(certificate);
        var signatureFactory = new Org.BouncyCastle.Crypto.Operators.Asn1SignatureFactory(
            "SHA256WITHRSA",
            Org.BouncyCastle.Security.DotNetUtilities.GetRsaKeyPair(rsa).Private);
        Org.BouncyCastle.Cms.SignerInfoGenerator signer =
            new Org.BouncyCastle.Cms.SignerInfoGeneratorBuilder().Build(signatureFactory, bcCertificate);
        var generator = new Org.BouncyCastle.Tsp.TimeStampTokenGenerator(
            signer,
            Org.BouncyCastle.Crypto.Operators.Asn1DigestFactory.Get("SHA256"),
            new Org.BouncyCastle.Asn1.DerObjectIdentifier("1.3.6.1.4.1.59069.1.1"),
            isIssuerSerialIncluded: true);
        generator.SetCertificates(new SingleCertificateStore(bcCertificate));
        var requestGenerator = new Org.BouncyCastle.Tsp.TimeStampRequestGenerator();
        requestGenerator.SetCertReq(true);
        Org.BouncyCastle.Tsp.TimeStampRequest request = requestGenerator.Generate(
            Org.BouncyCastle.Tsp.TspAlgorithms.Sha256,
            Org.BouncyCastle.Security.DigestUtilities.CalculateDigest("SHA256", expectedData));
        return generator.Generate(request, Org.BouncyCastle.Math.BigInteger.One, DateTime.UtcNow).GetEncoded();
    }

    private static X509Certificate2 CreateSigningCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=OfficeIMO PDF Signature Test",
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
        request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, critical: false));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }

    private static X509Certificate2 CreateEcdsaSigningCertificate() {
        using ECDsa ecdsa = ECDsa.Create(ECCurve.NamedCurves.nistP256);
        var request = new CertificateRequest("CN=OfficeIMO PDF ECDSA Test", ecdsa, HashAlgorithmName.SHA256);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
        request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, critical: false));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }

    private static X509Certificate2 CreateTimestampCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=OfficeIMO PDF Test TSA",
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
        request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
            new OidCollection { new Oid("1.3.6.1.5.5.7.3.8") },
            critical: true));
        request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, critical: false));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }

    private static PdfCmsSignatureCryptographyProvider CreateProvider(
        Func<X509Certificate2, X509Chain, bool> chainEvaluator) {
        var options = new CmsVerificationOptions();
        options.CertificateValidation.ChainEvaluator = chainEvaluator;
        return new PdfCmsSignatureCryptographyProvider(options);
    }

    private sealed class SingleCertificateStore :
        Org.BouncyCastle.Utilities.Collections.IStore<Org.BouncyCastle.X509.X509Certificate> {
        private readonly Org.BouncyCastle.X509.X509Certificate _certificate;

        internal SingleCertificateStore(Org.BouncyCastle.X509.X509Certificate certificate) {
            _certificate = certificate;
        }

        public IEnumerable<Org.BouncyCastle.X509.X509Certificate> EnumerateMatches(
            Org.BouncyCastle.Utilities.Collections.ISelector<Org.BouncyCastle.X509.X509Certificate> selector) {
            if (selector == null || selector.Match(_certificate)) yield return _certificate;
        }
    }

}
#endif
