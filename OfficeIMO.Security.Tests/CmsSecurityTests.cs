namespace OfficeIMO.Security.Tests;

public sealed class CmsSecurityTests {
    [Fact]
    public void DetachedSignature_RoundTrips_AndDetectsTampering() {
        byte[] content = Encoding.UTF8.GetBytes("OfficeIMO detached content\r\n");
        using X509Certificate2 certificate = CreateRsaCertificate("OfficeIMO CMS Signer");
        DateTimeOffset signingTime = DateTimeOffset.UtcNow;
        byte[] encoded = CmsSignedDataSigner.SignDetached(
            content,
            certificate,
            new CmsSigningOptions { SigningTime = signingTime });

        CmsVerificationResult valid = CmsSignedDataVerifier.VerifyDetached(encoded, content, TrustSelfSigned());
        byte[] tamperedContent = Encoding.UTF8.GetBytes("OfficeIMO tampered content\r\n");
        CmsVerificationResult tampered = CmsSignedDataVerifier.VerifyDetached(encoded, tamperedContent, TrustSelfSigned());

        Assert.True(valid.Parsed);
        Assert.True(valid.IsDetached);
        Assert.True(
            valid.IsCryptographicallyValid,
            string.Join(" | ", valid.Signers.SelectMany(static signer => signer.Findings)
                .Concat(valid.Findings)
                .Select(static finding => finding.Code + ": " + finding.Message)));
        Assert.Single(valid.Signers);
        Assert.Equal(SecurityValidationStatus.Valid, valid.Signers[0].CertificateValidation.ChainStatus);
        Assert.Equal(signingTime.ToUnixTimeSeconds(), valid.Signers[0].SigningTime?.ToUnixTimeSeconds());
        Assert.False(tampered.IsCryptographicallyValid);
        Assert.Equal(SecurityValidationStatus.Invalid, tampered.Signers[0].DigestStatus);
        Assert.Contains(tampered.Signers[0].Findings, finding => finding.Code == "CmsContentDigestMismatch");
    }

    [Fact]
    public void EncapsulatedSignature_ReturnsTheExactContent() {
        byte[] content = { 0, 1, 2, 3, 254, 255 };
        using X509Certificate2 certificate = CreateRsaCertificate("OfficeIMO CMS Encapsulated");
        byte[] encoded = CmsSignedDataSigner.SignEncapsulated(content, certificate);

        CmsVerificationResult result = CmsSignedDataVerifier.Verify(encoded, TrustSelfSigned());

        Assert.True(result.IsCryptographicallyValid);
        Assert.False(result.IsDetached);
        Assert.Equal(content, result.EncapsulatedContent);
    }

    [Fact]
    public void DetachedSignature_WithoutContent_IsIndeterminateAndActionable() {
        byte[] content = Encoding.ASCII.GetBytes("detached");
        using X509Certificate2 certificate = CreateRsaCertificate("OfficeIMO Detached Missing");
        byte[] encoded = CmsSignedDataSigner.SignDetached(content, certificate);

        CmsVerificationResult result = CmsSignedDataVerifier.Verify(encoded, TrustSelfSigned());

        Assert.True(result.Parsed);
        Assert.True(result.IsDetached);
        Assert.False(result.IsCryptographicallyValid);
        Assert.Equal(SecurityValidationStatus.Indeterminate, result.Signers[0].SignatureStatus);
        Assert.Contains(result.Findings, finding => finding.Code == "DetachedContentMissing");
    }

    [Fact]
    public void Envelope_RoundTripsForMatchingRecipient() {
        byte[] content = Encoding.UTF8.GetBytes("confidential OfficeIMO payload");
        using X509Certificate2 recipient = CreateRsaCertificate("OfficeIMO CMS Recipient");
        byte[] encoded = CmsEnvelopedDataService.Encrypt(content, new[] { recipient });

        CmsDecryptionResult result = CmsEnvelopedDataService.Decrypt(encoded, recipient);

        Assert.True(result.Parsed);
        Assert.True(result.Decrypted);
        Assert.Equal(content, result.Content);
        Assert.NotNull(result.ContentEncryptionAlgorithmOid);
        Assert.NotNull(result.KeyEncryptionAlgorithmOid);
    }

    [Fact]
    public void Envelope_ReportsNonMatchingRecipientWithoutThrowing() {
        byte[] content = Encoding.UTF8.GetBytes("confidential OfficeIMO payload");
        using X509Certificate2 recipient = CreateRsaCertificate("OfficeIMO CMS Recipient");
        using X509Certificate2 other = CreateRsaCertificate("OfficeIMO Other Recipient");
        byte[] encoded = CmsEnvelopedDataService.Encrypt(content, new[] { recipient });

        CmsDecryptionResult result = CmsEnvelopedDataService.Decrypt(encoded, other);

        Assert.True(result.Parsed);
        Assert.False(result.Decrypted);
        Assert.Contains(result.Findings, finding => finding.Code == "EnvelopeRecipientNotFound");
    }

    [Fact]
    public void Verification_EnforcesEncodedSizeLimitBeforeParsing() {
        var options = new CmsVerificationOptions { MaxEncodedBytes = 2 };

        ArgumentException exception = Assert.Throws<ArgumentException>(
            () => CmsSignedDataVerifier.Verify(new byte[] { 1, 2, 3 }, options));

        Assert.Contains("exceeds the configured limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Verification_AcceptsEcdsaCmsProducedByAnIndependentGenerator() {
        byte[] content = Encoding.UTF8.GetBytes("ECDSA interoperability");
        using X509Certificate2 certificate = CreateEcdsaCertificate("OfficeIMO ECDSA Signer");
        using ECDsa ecdsa = certificate.GetECDsaPrivateKey() ?? throw new InvalidOperationException();
        Org.BouncyCastle.X509.X509Certificate bcCertificate =
            Org.BouncyCastle.Security.DotNetUtilities.FromX509Certificate(certificate);
        Org.BouncyCastle.Crypto.AsymmetricKeyParameter privateKey =
            Org.BouncyCastle.Security.DotNetUtilities.GetECDsaKeyPair(ecdsa).Private;
        var generator = new Org.BouncyCastle.Cms.CmsSignedDataGenerator { UseDefiniteLength = true };
        var signatureFactory = new Org.BouncyCastle.Crypto.Operators.Asn1SignatureFactory(
            "SHA256WITHECDSA",
            privateKey);
        generator.AddSignerInfoGenerator(
            new Org.BouncyCastle.Cms.SignerInfoGeneratorBuilder().Build(signatureFactory, bcCertificate));
        generator.AddCertificate(bcCertificate);
        byte[] encoded = generator.Generate(
            new Org.BouncyCastle.Cms.CmsProcessableByteArray(content),
            encapsulate: true).GetEncoded();

        CmsVerificationResult result = CmsSignedDataVerifier.Verify(encoded, TrustSelfSigned());

        Assert.True(result.IsCryptographicallyValid);
        Assert.Equal("1.2.840.10045.4.3.2", result.Signers[0].SignatureAlgorithmOid);
    }

    [Fact]
    public void TimestampVerifier_ValidatesSignatureProfileAndMessageImprint() {
        byte[] timestampedData = Encoding.UTF8.GetBytes("PDF signature bytes");
        using X509Certificate2 certificate = CreateTimestampCertificate();
        using RSA rsa = certificate.GetRSAPrivateKey() ?? throw new InvalidOperationException();
        Org.BouncyCastle.X509.X509Certificate bcCertificate =
            Org.BouncyCastle.Security.DotNetUtilities.FromX509Certificate(certificate);
        Org.BouncyCastle.Crypto.AsymmetricKeyParameter privateKey =
            Org.BouncyCastle.Security.DotNetUtilities.GetRsaKeyPair(rsa).Private;
        var signerFactory = new Org.BouncyCastle.Crypto.Operators.Asn1SignatureFactory(
            "SHA256WITHRSA",
            privateKey);
        Org.BouncyCastle.Cms.SignerInfoGenerator signer =
            new Org.BouncyCastle.Cms.SignerInfoGeneratorBuilder().Build(signerFactory, bcCertificate);
        var generator = new Org.BouncyCastle.Tsp.TimeStampTokenGenerator(
            signer,
            Org.BouncyCastle.Crypto.Operators.Asn1DigestFactory.Get("SHA256"),
            new Org.BouncyCastle.Asn1.DerObjectIdentifier("1.3.6.1.4.1.59069.1.1"),
            isIssuerSerialIncluded: true);
        generator.SetCertificates(new SingleCertificateStore(bcCertificate));
        var requestGenerator = new Org.BouncyCastle.Tsp.TimeStampRequestGenerator();
        requestGenerator.SetCertReq(true);
        byte[] imprint = Org.BouncyCastle.Security.DigestUtilities.CalculateDigest("SHA256", timestampedData);
        Org.BouncyCastle.Tsp.TimeStampRequest request = requestGenerator.Generate(
            Org.BouncyCastle.Tsp.TspAlgorithms.Sha256,
            imprint);
        byte[] encoded = generator.Generate(
            request,
            Org.BouncyCastle.Math.BigInteger.One,
            DateTime.UtcNow).GetEncoded();
        var trust = new CertificateValidationOptions { ChainEvaluator = static (_, _) => true };

        Rfc3161TimestampVerificationResult valid = Rfc3161TimestampVerifier.Verify(encoded, timestampedData, trust);
        Rfc3161TimestampVerificationResult tampered = Rfc3161TimestampVerifier.Verify(
            encoded,
            Encoding.UTF8.GetBytes("different signature bytes"),
            trust);

        Assert.Equal(SecurityValidationStatus.Valid, valid.Status);
        Assert.Equal(SecurityValidationStatus.Valid, valid.CertificateValidation.ChainStatus);
        Assert.NotNull(valid.Timestamp);
        Assert.Equal("2.16.840.1.101.3.4.2.1", valid.MessageImprintAlgorithmOid);
        Assert.Equal(SecurityValidationStatus.Invalid, tampered.Status);
        Assert.Contains(tampered.Findings, finding => finding.Code == "TimestampImprintMismatch");
    }

    private static CmsVerificationOptions TrustSelfSigned() {
        var options = new CmsVerificationOptions();
        options.CertificateValidation.ChainEvaluator = static (_, _) => true;
        return options;
    }

    private static X509Certificate2 CreateRsaCertificate(string commonName) {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=" + commonName,
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(
            X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment,
            critical: true));
        request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, critical: false));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }

    private static X509Certificate2 CreateEcdsaCertificate(string commonName) {
        using ECDsa ecdsa = ECDsa.Create(ECCurve.NamedCurves.nistP256);
        var request = new CertificateRequest("CN=" + commonName, ecdsa, HashAlgorithmName.SHA256);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
        request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, critical: false));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }

    private static X509Certificate2 CreateTimestampCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=OfficeIMO Test TSA",
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
