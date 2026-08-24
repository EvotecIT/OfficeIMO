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

    [Theory]
    [InlineData("SHA1", "1.3.14.3.2.26")]
    [InlineData("SHA256", "2.16.840.1.101.3.4.2.1")]
    [InlineData("SHA384", "2.16.840.1.101.3.4.2.2")]
    [InlineData("SHA512", "2.16.840.1.101.3.4.2.3")]
    public void DetachedRsaSignature_FastPathPreservesSupportedDigestContracts(
        string digestName,
        string expectedDigestOid) {
        byte[] content = Encoding.UTF8.GetBytes("OfficeIMO RSA digest contract\r\n");
        using X509Certificate2 certificate = CreateRsaCertificate("OfficeIMO CMS Digest " + digestName);
        byte[] encoded = CmsSignedDataSigner.SignDetached(
            content,
            certificate,
            new CmsSigningOptions {
                DigestAlgorithm = new HashAlgorithmName(digestName),
                IncludeSigningTime = false,
                IncludeCertificateChain = false
            });

        CmsVerificationResult result = CmsSignedDataVerifier.VerifyDetached(encoded, content, TrustSelfSigned());

        CmsSignerVerificationResult signer = Assert.Single(result.Signers);
        Assert.True(result.IsCryptographicallyValid);
        Assert.Equal(SecurityValidationStatus.Valid, signer.SignatureStatus);
        Assert.Equal(SecurityValidationStatus.Valid, signer.DigestStatus);
        Assert.Equal(expectedDigestOid, signer.DigestAlgorithmOid);
    }

    [Fact]
    public void PlatformDetachedRsaSignature_PreservesTypedResultAndRejectsTampering() {
        byte[] content = Encoding.UTF8.GetBytes("OfficeIMO platform CMS contract\r\n");
        using X509Certificate2 certificate = CreateRsaCertificate("OfficeIMO Platform CMS");
        var signedCms = new System.Security.Cryptography.Pkcs.SignedCms(
            new System.Security.Cryptography.Pkcs.ContentInfo(content),
            detached: true);
        var platformSigner = new System.Security.Cryptography.Pkcs.CmsSigner(
            System.Security.Cryptography.Pkcs.SubjectIdentifierType.IssuerAndSerialNumber,
            certificate) {
            DigestAlgorithm = new Oid("2.16.840.1.101.3.4.2.1"),
            IncludeOption = X509IncludeOption.EndCertOnly
        };
        signedCms.ComputeSignature(platformSigner, silent: true);
        byte[] encoded = signedCms.Encode();
        var options = new CmsVerificationOptions { ValidateTimestamps = false };
        options.CertificateValidation.ValidateChain = false;
        options.CertificateValidation.DisableCertificateDownloads = true;
        options.CertificateValidation.RevocationMode = X509RevocationMode.NoCheck;

        CmsVerificationResult valid = CmsSignedDataVerifier.VerifyDetached(encoded, content, options);
        byte[] tampered = (byte[])content.Clone();
        tampered[^1] ^= 0x5A;
        CmsVerificationResult invalid = CmsSignedDataVerifier.VerifyDetached(encoded, tampered, options);

        CmsSignerVerificationResult signer = Assert.Single(valid.Signers);
        Assert.True(valid.Parsed);
        Assert.True(valid.IsDetached);
        Assert.True(valid.IsCryptographicallyValid);
        Assert.Equal(SecurityValidationStatus.Valid, signer.SignatureStatus);
        Assert.Equal(SecurityValidationStatus.Valid, signer.DigestStatus);
        Assert.Equal(SecurityValidationStatus.NotPerformed, signer.CertificateValidation.ChainStatus);
        Assert.Equal(certificate.RawData, signer.SignerCertificate);
        Assert.Equal(certificate.Subject, signer.Subject);
        Assert.Equal(certificate.Issuer, signer.Issuer);
        Assert.Equal(certificate.SerialNumber, signer.SerialNumber);
        Assert.Equal(certificate.Thumbprint, signer.Thumbprint);
        Assert.False(invalid.IsCryptographicallyValid);
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
    public void EncapsulatedSignature_PreservesACallerSelectedContentType() {
        const string contentTypeOid = "1.3.6.1.4.1.311.2.1.4";
        byte[] content = { 48, 0 };
        using X509Certificate2 certificate = CreateRsaCertificate("OfficeIMO CMS Custom Content Type");
        byte[] encoded = CmsSignedDataSigner.SignEncapsulated(
            content,
            certificate,
            new CmsSigningOptions { ContentTypeOid = contentTypeOid });

        CmsVerificationResult result = CmsSignedDataVerifier.Verify(encoded, TrustSelfSigned());

        Assert.True(result.IsCryptographicallyValid);
        Assert.Equal(contentTypeOid, result.ContentTypeOid);
        Assert.Equal(content, result.EncapsulatedContent);
    }

    [Fact]
    public void EncapsulatedSignature_RejectsAnInvalidContentTypeOid() {
        using X509Certificate2 certificate = CreateRsaCertificate("OfficeIMO CMS Invalid Content Type");

        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            CmsSignedDataSigner.SignEncapsulated(
                new byte[] { 1 },
                certificate,
                new CmsSigningOptions { ContentTypeOid = "not-an-oid" }));

        Assert.Equal("options", exception.ParamName);
    }

    [Fact]
    public void Verification_RejectsTlsOnlySignerCertificates() {
        byte[] content = Encoding.UTF8.GetBytes("TLS certificates are not document signers");
        using X509Certificate2 certificate = CreateRsaCertificate(
            "OfficeIMO TLS Only",
            new Oid("1.3.6.1.5.5.7.3.1"));
        byte[] encoded = CmsSignedDataSigner.SignEncapsulated(content, certificate);

        CmsVerificationResult result = CmsSignedDataVerifier.Verify(encoded, TrustSelfSigned());

        CmsSignerVerificationResult signer = Assert.Single(result.Signers);
        Assert.Equal(SecurityValidationStatus.Invalid, signer.CertificateValidation.ChainStatus);
        Assert.Contains(signer.Findings, finding => finding.Code == "CertificateEnhancedKeyUsageInvalid");
    }

    [Fact]
    public void DocumentSigningValidationRejectsEmailProtectionOnlyCertificates() {
        using X509Certificate2 certificate = CreateRsaCertificate(
            "OfficeIMO Email Protection Only",
            new Oid("1.3.6.1.5.5.7.3.4"));
        var options = new CertificateValidationOptions {
            ChainEvaluator = static (_, _) => true,
            RevocationMode = X509RevocationMode.NoCheck
        };

        CertificateTrustValidationResult result = CertificateValidator.Validate(
            certificate,
            options: options,
            purpose: CertificateValidationPurpose.DocumentSigning);

        Assert.Equal(SecurityValidationStatus.Invalid, result.Validation.ChainStatus);
        Assert.Contains(result.Findings, finding => finding.Code == "CertificateEnhancedKeyUsageInvalid");
    }

    [Fact]
    public void EmailSigningValidationAcceptsEmailProtectionCertificates() {
        using X509Certificate2 certificate = CreateRsaCertificate(
            "OfficeIMO Email Protection",
            new Oid("1.3.6.1.5.5.7.3.4"));
        var options = new CertificateValidationOptions {
            ChainEvaluator = static (_, _) => true,
            RevocationMode = X509RevocationMode.NoCheck
        };

        CertificateTrustValidationResult result = CertificateValidator.Validate(
            certificate,
            options: options,
            purpose: CertificateValidationPurpose.EmailSigning);

        Assert.Equal(SecurityValidationStatus.Valid, result.Validation.ChainStatus);
        Assert.DoesNotContain(result.Findings, finding => finding.Code == "CertificateEnhancedKeyUsageInvalid");
    }

    [Theory]
    [InlineData("1.3.6.1.5.5.7.3.3")]
    [InlineData("1.3.6.1.5.5.7.3.36")]
    [InlineData("1.3.6.1.4.1.311.10.3.12")]
    public void EmailSigningValidationRejectsNonEmailSigningCertificates(string enhancedKeyUsageOid) {
        using X509Certificate2 certificate = CreateRsaCertificate(
            "OfficeIMO Non-Email Signer",
            new Oid(enhancedKeyUsageOid));
        var options = new CertificateValidationOptions {
            ChainEvaluator = static (_, _) => true,
            RevocationMode = X509RevocationMode.NoCheck
        };

        CertificateTrustValidationResult result = CertificateValidator.Validate(
            certificate,
            options: options,
            purpose: CertificateValidationPurpose.EmailSigning);

        Assert.Equal(SecurityValidationStatus.Invalid, result.Validation.ChainStatus);
        Assert.Contains(result.Findings, finding => finding.Code == "CertificateEnhancedKeyUsageInvalid");
    }

    [Fact]
    public void EmailSigningValidationAcceptsAnyExtendedKeyUsageCertificates() {
        using X509Certificate2 certificate = CreateRsaCertificate(
            "OfficeIMO Any Extended Key Usage",
            new Oid("2.5.29.37.0"));
        var options = new CertificateValidationOptions {
            ChainEvaluator = static (_, _) => true,
            RevocationMode = X509RevocationMode.NoCheck
        };

        CertificateTrustValidationResult result = CertificateValidator.Validate(
            certificate,
            options: options,
            purpose: CertificateValidationPurpose.EmailSigning);

        Assert.Equal(SecurityValidationStatus.Valid, result.Validation.ChainStatus);
        Assert.DoesNotContain(result.Findings, finding => finding.Code == "CertificateEnhancedKeyUsageInvalid");
    }

    [Fact]
    public void CertificateUsageValidationRemainsActiveWhenPlatformChainBuildingIsDisabled() {
        using X509Certificate2 certificate = CreateRsaCertificate("OfficeIMO Non-Timestamp Authority");
        using X509Certificate2 timestampAuthority = CreateTimestampCertificate();
        var options = new CertificateValidationOptions { ValidateChain = false };

        CertificateTrustValidationResult invalid = CertificateValidator.Validate(
            certificate,
            options: options,
            purpose: CertificateValidationPurpose.TimestampAuthority);
        CertificateTrustValidationResult valid = CertificateValidator.Validate(
            timestampAuthority,
            options: options,
            purpose: CertificateValidationPurpose.TimestampAuthority);

        Assert.Equal(SecurityValidationStatus.Invalid, invalid.Validation.ChainStatus);
        Assert.Contains(invalid.Findings, finding => finding.Code == "CertificateEnhancedKeyUsageInvalid");
        Assert.Equal(SecurityValidationStatus.NotPerformed, valid.Validation.ChainStatus);
        Assert.DoesNotContain(valid.Findings, finding => finding.Code is
            "CertificateKeyUsageInvalid" or "CertificateEnhancedKeyUsageInvalid");
    }

    [Fact]
    public void CertificateValidationRecognizesACompleteOfflineIssuerPath() {
        using RSA rootKey = RSA.Create(2048);
        var rootRequest = new CertificateRequest(
            "CN=OfficeIMO Offline Root",
            rootKey,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        rootRequest.CertificateExtensions.Add(new X509BasicConstraintsExtension(true, false, 0, true));
        rootRequest.CertificateExtensions.Add(new X509KeyUsageExtension(
            X509KeyUsageFlags.KeyCertSign | X509KeyUsageFlags.CrlSign,
            true));
        rootRequest.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(rootRequest.PublicKey, false));
        using X509Certificate2 root = rootRequest.CreateSelfSigned(
            DateTimeOffset.UtcNow.AddMinutes(-5),
            DateTimeOffset.UtcNow.AddDays(1));

        using RSA leafKey = RSA.Create(2048);
        var leafRequest = new CertificateRequest(
            "CN=OfficeIMO Offline Leaf",
            leafKey,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        leafRequest.CertificateExtensions.Add(new X509BasicConstraintsExtension(false, false, 0, true));
        leafRequest.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        leafRequest.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(leafRequest.PublicKey, false));
        using X509Certificate2 leaf = leafRequest.Create(
            root,
            DateTimeOffset.UtcNow.AddMinutes(-5),
            DateTimeOffset.UtcNow.AddHours(12),
            new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 });

        Assert.False(CertificateChainValidator.HasCompleteOfflinePath(
            leaf,
            Array.Empty<X509Certificate2>()));
        Assert.True(CertificateChainValidator.HasCompleteOfflinePath(
            leaf,
            new[] { root }));
    }

    [Fact]
    public void CertificateOfflineIssuerPathSearchStopsAtTheCryptographicWorkLimit() {
        using RSA key = RSA.Create(2048);
        X509SignatureGenerator generator = X509SignatureGenerator.CreateForRSA(key, RSASignaturePadding.Pkcs1);
        DateTimeOffset notBefore = DateTimeOffset.UtcNow.AddMinutes(-5);
        DateTimeOffset notAfter = DateTimeOffset.UtcNow.AddDays(1);

        using X509Certificate2 issuerOne = CreateOfflinePathCertificate(
            "CN=OfficeIMO Alternate Issuer",
            "CN=OfficeIMO Missing Root",
            key,
            generator,
            notBefore,
            notAfter,
            1);
        using X509Certificate2 issuerTwo = CreateOfflinePathCertificate(
            "CN=OfficeIMO Alternate Issuer",
            "CN=OfficeIMO Missing Root",
            key,
            generator,
            notBefore,
            notAfter,
            2);
        using X509Certificate2 leaf = CreateOfflinePathCertificate(
            "CN=OfficeIMO Offline Search Leaf",
            "CN=OfficeIMO Alternate Issuer",
            key,
            generator,
            notBefore,
            notAfter,
            3);

        OfflineCertificatePathSearchOutcome bounded = CertificateChainValidator.FindCompleteOfflinePath(
            leaf,
            new[] { issuerOne, issuerTwo },
            maxIssuerSignatureChecks: 1);
        OfflineCertificatePathSearchOutcome completed = CertificateChainValidator.FindCompleteOfflinePath(
            leaf,
            new[] { issuerOne, issuerTwo },
            maxIssuerSignatureChecks: 2);

        Assert.Equal(OfflineCertificatePathSearchOutcome.WorkLimitExceeded, bounded);
        Assert.Equal(OfflineCertificatePathSearchOutcome.Incomplete, completed);
    }

    [Fact]
    public void EncapsulatedSignature_StopsAtTheConfiguredContentLimit() {
        byte[] content = Enumerable.Repeat((byte)0x5a, 4096).ToArray();
        using X509Certificate2 certificate = CreateRsaCertificate("OfficeIMO CMS Bounded Encapsulated");
        byte[] encoded = CmsSignedDataSigner.SignEncapsulated(content, certificate);
        CmsVerificationOptions options = TrustSelfSigned();
        options.MaxContentBytes = 32;

        CmsVerificationResult result = CmsSignedDataVerifier.Verify(encoded, options);

        Assert.True(result.Parsed);
        Assert.Null(result.EncapsulatedContent);
        Assert.Contains(result.Findings, finding => finding.Code == "CmsContentLimitExceeded" &&
            finding.Message.Contains("configured limit of 32 bytes", StringComparison.Ordinal));
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
    public void Envelope_StopsDecryptionAtTheConfiguredContentLimit() {
        byte[] content = Enumerable.Repeat((byte)0xa5, 4096).ToArray();
        using X509Certificate2 recipient = CreateRsaCertificate("OfficeIMO CMS Bounded Recipient");
        byte[] encoded = CmsEnvelopedDataService.Encrypt(content, new[] { recipient });
        var options = new CmsEnvelopeOptions { MaxContentBytes = 32 };

        CmsDecryptionResult result = CmsEnvelopedDataService.Decrypt(encoded, recipient, options);

        Assert.True(result.Parsed);
        Assert.False(result.Decrypted);
        Assert.Null(result.Content);
        Assert.Contains(result.Findings, finding => finding.Code == "EnvelopeContentLimitExceeded" &&
            finding.Message.Contains("configured limit of 32 bytes", StringComparison.Ordinal));
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
    public void Verification_BoundsTimestampTokensAcrossTheCmsOperation() {
        byte[] content = Encoding.UTF8.GetBytes("CMS timestamp budget");
        using X509Certificate2 certificate = CreateRsaCertificate("OfficeIMO Timestamp Budget");
        byte[] encoded = CmsSignedDataSigner.SignEncapsulated(content, certificate);
        var signedData = new Org.BouncyCastle.Cms.CmsSignedData(encoded);
        Org.BouncyCastle.Cms.SignerInformation signer =
            signedData.GetSignerInfos().GetSigners().Single();
        var timestampValues = new Org.BouncyCastle.Asn1.Asn1EncodableVector();
        timestampValues.Add(new Org.BouncyCastle.Asn1.DerSequence());
        timestampValues.Add(new Org.BouncyCastle.Asn1.DerSequence());
        timestampValues.Add(new Org.BouncyCastle.Asn1.DerSequence());
        Org.BouncyCastle.Asn1.DerObjectIdentifier timestampOid =
            Org.BouncyCastle.Asn1.Pkcs.PkcsObjectIdentifiers.IdAASignatureTimeStampToken;
        var timestampAttribute = new Org.BouncyCastle.Asn1.Cms.Attribute(
            timestampOid,
            new Org.BouncyCastle.Asn1.DerSet(timestampValues));
        var unsignedAttributes = new Org.BouncyCastle.Asn1.Cms.AttributeTable(
            new Dictionary<Org.BouncyCastle.Asn1.DerObjectIdentifier, object> {
                [timestampOid] = timestampAttribute
            });
        Org.BouncyCastle.Cms.SignerInformation withTimestamps =
            Org.BouncyCastle.Cms.SignerInformation.ReplaceUnsignedAttributes(signer, unsignedAttributes);
        Org.BouncyCastle.Cms.CmsSignedData repeated = Org.BouncyCastle.Cms.CmsSignedData.ReplaceSigners(
            signedData,
            new Org.BouncyCastle.Cms.SignerInformationStore(new[] { withTimestamps }));
        CmsVerificationOptions options = TrustSelfSigned();
        options.MaxTimestampTokens = 2;

        CmsVerificationResult result = CmsSignedDataVerifier.Verify(repeated.GetEncoded(), options);

        CmsSignerVerificationResult verifiedSigner = Assert.Single(result.Signers);
        Assert.Equal(SecurityValidationStatus.Invalid, verifiedSigner.TimestampStatus);
        Assert.Contains(verifiedSigner.Findings,
            finding => finding.Code == "CmsTimestampCountLimitExceeded");
    }

    [Fact]
    public void Verification_BoundsTimestampTokenWhileEncodingUnsignedAttributes() {
        byte[] content = Encoding.UTF8.GetBytes("CMS timestamp encoding budget");
        using X509Certificate2 certificate = CreateRsaCertificate("OfficeIMO Timestamp Encoding Budget");
        byte[] encoded = CmsSignedDataSigner.SignEncapsulated(content, certificate);
        var signedData = new Org.BouncyCastle.Cms.CmsSignedData(encoded);
        Org.BouncyCastle.Cms.SignerInformation signer =
            signedData.GetSignerInfos().GetSigners().Single();
        Org.BouncyCastle.Asn1.DerObjectIdentifier timestampOid =
            Org.BouncyCastle.Asn1.Pkcs.PkcsObjectIdentifiers.IdAASignatureTimeStampToken;
        var timestampAttribute = new Org.BouncyCastle.Asn1.Cms.Attribute(
            timestampOid,
            new Org.BouncyCastle.Asn1.DerSet(
                new Org.BouncyCastle.Asn1.DerOctetString(new byte[4096])));
        var unsignedAttributes = new Org.BouncyCastle.Asn1.Cms.AttributeTable(
            new Dictionary<Org.BouncyCastle.Asn1.DerObjectIdentifier, object> {
                [timestampOid] = timestampAttribute
            });
        Org.BouncyCastle.Cms.SignerInformation withTimestamp =
            Org.BouncyCastle.Cms.SignerInformation.ReplaceUnsignedAttributes(signer, unsignedAttributes);
        Org.BouncyCastle.Cms.CmsSignedData oversized = Org.BouncyCastle.Cms.CmsSignedData.ReplaceSigners(
            signedData,
            new Org.BouncyCastle.Cms.SignerInformationStore(new[] { withTimestamp }));
        CmsVerificationOptions options = TrustSelfSigned();
        options.MaxTimestampTokenBytes = 32;

        CmsVerificationResult result = CmsSignedDataVerifier.Verify(oversized.GetEncoded(), options);

        CmsSignerVerificationResult verifiedSigner = Assert.Single(result.Signers);
        Assert.Equal(SecurityValidationStatus.Invalid, verifiedSigner.TimestampStatus);
        Assert.Contains(verifiedSigner.Findings,
            finding => finding.Code == "CmsTimestampSizeLimitExceeded");
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
        (byte[] encoded, DateTime generationTime) = CreateTimestampToken(timestampedData, certificate);
        DateTime? observedDefaultVerificationTime = null;
        var trust = new CertificateValidationOptions {
            ChainEvaluator = (_, chain) => {
                observedDefaultVerificationTime = chain.ChainPolicy.VerificationTime;
                return true;
            }
        };

        Rfc3161TimestampVerificationResult valid = Rfc3161TimestampVerifier.Verify(encoded, timestampedData, trust);
        Rfc3161TimestampVerificationResult tampered = Rfc3161TimestampVerifier.Verify(
            encoded,
            Encoding.UTF8.GetBytes("different signature bytes"),
            trust);
        Rfc3161TimestampVerificationResult untrusted = Rfc3161TimestampVerifier.Verify(
            encoded,
            timestampedData,
            new CertificateValidationOptions { ChainEvaluator = static (_, _) => false });
        DateTime explicitVerificationTime = generationTime.AddSeconds(30);
        DateTime? observedExplicitVerificationTime = null;
        var explicitTrust = new CertificateValidationOptions {
            VerificationTime = explicitVerificationTime,
            ChainEvaluator = (_, chain) => {
                observedExplicitVerificationTime = chain.ChainPolicy.VerificationTime;
                return true;
            }
        };
        Rfc3161TimestampVerifier.Verify(encoded, timestampedData, explicitTrust);

        Assert.Equal(SecurityValidationStatus.Valid, valid.Status);
        Assert.Equal(SecurityValidationStatus.Valid, valid.CertificateValidation.ChainStatus);
        Assert.NotNull(valid.Timestamp);
        Assert.Equal(valid.Timestamp.Value.UtcDateTime, observedDefaultVerificationTime);
        Assert.Equal(explicitVerificationTime, observedExplicitVerificationTime);
        Assert.Null(trust.VerificationTime);
        Assert.Equal("2.16.840.1.101.3.4.2.1", valid.MessageImprintAlgorithmOid);
        Assert.Equal(SecurityValidationStatus.Invalid, tampered.Status);
        Assert.Contains(tampered.Findings, finding => finding.Code == "TimestampImprintMismatch");
        Assert.Equal(SecurityValidationStatus.Invalid, untrusted.Status);
        Assert.Equal(SecurityValidationStatus.Invalid, untrusted.CertificateValidation.ChainStatus);
    }

    [Fact]
    public void TimestampVerifier_ResolvesOmittedTsaCertificateFromCallerExtras() {
        byte[] timestampedData = Encoding.UTF8.GetBytes("timestamp with external TSA certificate");
        using X509Certificate2 certificate = CreateTimestampCertificate();
        (byte[] encoded, _) = CreateTimestampToken(
            timestampedData,
            certificate,
            includeCertificate: false);
        var trust = new CertificateValidationOptions {
            ChainEvaluator = static (_, _) => true
        };
        trust.ExtraCertificates.Add(certificate);

        Rfc3161TimestampVerificationResult result = Rfc3161TimestampVerifier.Verify(
            encoded,
            timestampedData,
            trust);

        Assert.Equal(SecurityValidationStatus.Valid, result.Status);
        Assert.Equal(SecurityValidationStatus.Valid, result.CertificateValidation.ChainStatus);
        Assert.DoesNotContain(result.Findings, finding => finding.Code == "TimestampCertificateMissing");
        Assert.Equal(certificate.RawData, result.TsaCertificate);
    }

    [Theory]
    [InlineData(SecurityValidationStatus.Invalid, SecurityValidationStatus.Invalid)]
    [InlineData(SecurityValidationStatus.Indeterminate, SecurityValidationStatus.Indeterminate)]
    [InlineData(SecurityValidationStatus.Valid, SecurityValidationStatus.Valid)]
    [InlineData(SecurityValidationStatus.NotPerformed, SecurityValidationStatus.Valid)]
    public void TimestampVerifier_CombinesTsaRevocationWithTrust(
        SecurityValidationStatus revocationStatus,
        SecurityValidationStatus expectedStatus) {
        SecurityValidationStatus status = Rfc3161TimestampVerifier.ResolveTimestampStatus(
            signatureValid: true,
            imprintValid: true,
            certificateStatus: SecurityValidationStatus.Valid,
            revocationStatus);

        Assert.Equal(expectedStatus, status);
    }

    [Fact]
    public void CmsSignerTrustUsesEarliestValidTimestampUnlessCallerOverridesTime() {
        DateTimeOffset earlier = DateTimeOffset.UtcNow.AddYears(-2);
        DateTimeOffset later = earlier.AddMinutes(5);
        var timestamps = new[] {
            CreateTimestampResult(SecurityValidationStatus.Valid, later),
            CreateTimestampResult(SecurityValidationStatus.Valid, earlier),
            CreateTimestampResult(SecurityValidationStatus.Invalid, earlier.AddYears(-1))
        };
        var source = new CertificateValidationOptions();

        CertificateValidationOptions resolved =
            CmsSignedDataVerifier.ResolveSignerCertificateValidation(source, timestamps);

        Assert.Equal(earlier.UtcDateTime, resolved.VerificationTime);
        Assert.Null(source.VerificationTime);

        DateTime explicitTime = DateTime.UtcNow.AddDays(-3);
        source.VerificationTime = explicitTime;
        CertificateValidationOptions explicitResult =
            CmsSignedDataVerifier.ResolveSignerCertificateValidation(source, timestamps);
        Assert.Equal(explicitTime, explicitResult.VerificationTime);
    }

    private static Rfc3161TimestampVerificationResult CreateTimestampResult(
        SecurityValidationStatus status,
        DateTimeOffset timestamp) =>
        new Rfc3161TimestampVerificationResult(
            status,
            timestamp,
            policyOid: null,
            messageImprintAlgorithmOid: null,
            tsaCertificate: null,
            certificateValidation: new CertificateValidationResult(
                SecurityValidationStatus.Valid,
                SecurityValidationStatus.NotPerformed,
                Array.Empty<string>()),
            findings: Array.Empty<SecurityFinding>());

    private static CmsVerificationOptions TrustSelfSigned() {
        var options = new CmsVerificationOptions();
        options.CertificateValidation.ChainEvaluator = static (_, _) => true;
        return options;
    }

    private static X509Certificate2 CreateRsaCertificate(string commonName, Oid? enhancedKeyUsage = null) {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=" + commonName,
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(
            X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment,
            critical: true));
        if (enhancedKeyUsage != null) {
            request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
                new OidCollection { enhancedKeyUsage },
                critical: true));
        }
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

    private static X509Certificate2 CreateOfflinePathCertificate(
        string subject,
        string issuer,
        RSA key,
        X509SignatureGenerator generator,
        DateTimeOffset notBefore,
        DateTimeOffset notAfter,
        byte serial) {
        var request = new CertificateRequest(
            subject,
            key,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509BasicConstraintsExtension(true, false, 0, true));
        request.CertificateExtensions.Add(new X509KeyUsageExtension(
            X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyCertSign,
            true));
        return request.Create(
            new X500DistinguishedName(issuer),
            generator,
            notBefore,
            notAfter,
            new[] { serial });
    }

    private static (byte[] Encoded, DateTime GenerationTime) CreateTimestampToken(
        byte[] timestampedData,
        X509Certificate2 certificate,
        bool includeCertificate = true) {
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
        if (includeCertificate) {
            generator.SetCertificates(new SingleCertificateStore(bcCertificate));
        }
        var requestGenerator = new Org.BouncyCastle.Tsp.TimeStampRequestGenerator();
        requestGenerator.SetCertReq(includeCertificate);
        byte[] imprint = Org.BouncyCastle.Security.DigestUtilities.CalculateDigest("SHA256", timestampedData);
        Org.BouncyCastle.Tsp.TimeStampRequest request = requestGenerator.Generate(
            Org.BouncyCastle.Tsp.TspAlgorithms.Sha256,
            imprint);
        DateTime generationTime = DateTime.UtcNow.AddMinutes(-1);
        byte[] encoded = generator.Generate(
            request,
            Org.BouncyCastle.Math.BigInteger.One,
            generationTime).GetEncoded();
        return (encoded, generationTime);
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
