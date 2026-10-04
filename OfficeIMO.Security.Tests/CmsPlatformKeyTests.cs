using System.Security.Cryptography.Pkcs;
using Org.BouncyCastle.Cms;

namespace OfficeIMO.Security.Tests;

public sealed class CmsPlatformKeyTests {
    [Theory]
    [InlineData("2.16.840.1.101.3.4.1.2", false, false)]
    [InlineData("2.16.840.1.101.3.4.1.22", false, true)]
    [InlineData("2.16.840.1.101.3.4.1.42", true, false)]
    [InlineData("2.16.840.1.101.3.4.1.42", true, true)]
    public void DecryptsPlatformProducedEnvelopesWithoutExportingPrivateParameters(string algorithm, bool oaep, bool subjectKeyId) {
        using RSA key = RSA.Create(2048);
        var request = new CertificateRequest("CN=OfficeIMO Platform Recipient", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, false));
        using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1));
        byte[] content = Encoding.UTF8.GetBytes("Żółć 日本語 platform recipient");
        var platform = new EnvelopedCms(new ContentInfo(content), new AlgorithmIdentifier(new Oid(algorithm)));
        platform.Encrypt(new CmsRecipient(subjectKeyId ? SubjectIdentifierType.SubjectKeyIdentifier : SubjectIdentifierType.IssuerAndSerialNumber,
            certificate, oaep ? RSAEncryptionPadding.OaepSHA256 : RSAEncryptionPadding.Pkcs1));
        byte[] encoded = platform.Encode();
        var envelope = new CmsEnvelopedData(encoded);
        RecipientInformation recipient = Assert.Single(envelope.GetRecipientInfos().GetRecipients());
        using var nonExportable = new OperationOnlyRsa(key);
        Assert.Equal(content, PlatformCmsEnvelopeDecryptor.Decrypt(envelope, recipient, nonExportable, content.Length));
        Assert.True(nonExportable.DecryptCalls > 0);
        Assert.Equal(0, nonExportable.PrivateExportCalls);
        CmsDecryptionResult publicResult = CmsEnvelopedDataService.Decrypt(encoded, certificate);
        Assert.True(publicResult.Decrypted, string.Join("; ", publicResult.Findings.Select(finding => finding.Message)));
        Assert.Equal(content, publicResult.Content);
    }

    [Theory]
    [InlineData("RSA/NONE/OAEPWITHSHA256ANDMGF1WITHSHA1PADDING")]
    [InlineData("RSA/NONE/OAEPWITHSHA224ANDMGF1PADDING")]
    public void ExistingExportableAdapterKeepsOaepProfilesOutsideThePlatformContract(string algorithm) {
        using RSA key = RSA.Create(2048);
        var request = new CertificateRequest("CN=OfficeIMO OAEP Compatibility", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1));
        var generator = new CmsEnvelopedDataGenerator();
        var bcCertificate = Org.BouncyCastle.Security.DotNetUtilities.FromX509Certificate(certificate);
        generator.AddRecipientInfoGenerator(new KeyTransRecipientInfoGenerator(bcCertificate,
            new Org.BouncyCastle.Crypto.Operators.Asn1KeyWrapper(algorithm, bcCertificate)));
        byte[] content = Encoding.UTF8.GetBytes("OAEP compatibility");
        byte[] encoded = generator.Generate(new CmsProcessableByteArray(content), new Org.BouncyCastle.Asn1.DerObjectIdentifier(CmsEnvelopedGenerator.Aes256Cbc)).GetEncoded();
        var envelope = new CmsEnvelopedData(encoded);
        RecipientInformation recipient = Assert.Single(envelope.GetRecipientInfos().GetRecipients());
        Assert.False(PlatformCmsEnvelopeDecryptor.Supports(envelope, recipient));
        CmsDecryptionResult result = CmsEnvelopedDataService.Decrypt(encoded, certificate);
        Assert.True(result.Decrypted, string.Join("; ", result.Findings.Select(finding => finding.Message)));
        Assert.Equal(content, result.Content);
    }

    [Fact]
    public void RejectsOversizedPlaintextBeforeInvokingThePlatformKey() {
        using RSA key = RSA.Create(2048);
        var request = new CertificateRequest("CN=OfficeIMO Platform Bound", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1));
        var envelope = new CmsEnvelopedData(CmsEnvelopedDataService.Encrypt(new byte[4096], new[] { certificate }));
        using var nonExportable = new OperationOnlyRsa(key);
        Assert.Throws<SecurityContentLimitExceededException>(() => PlatformCmsEnvelopeDecryptor.Decrypt(envelope,
            Assert.Single(envelope.GetRecipientInfos().GetRecipients()), nonExportable, 32));
        Assert.Equal(0, nonExportable.DecryptCalls);
    }

    private sealed class OperationOnlyRsa : RSA {
        private readonly RSA _inner;
        internal OperationOnlyRsa(RSA inner) => _inner = inner;
        internal int DecryptCalls { get; private set; }
        internal int PrivateExportCalls { get; private set; }
        public override int KeySize { get => _inner.KeySize; set => throw new NotSupportedException(); }
        public override KeySizes[] LegalKeySizes => _inner.LegalKeySizes;
        public override byte[] Decrypt(byte[] data, RSAEncryptionPadding padding) { DecryptCalls++; return _inner.Decrypt(data, padding); }
        public override byte[] Encrypt(byte[] data, RSAEncryptionPadding padding) => _inner.Encrypt(data, padding);
        public override RSAParameters ExportParameters(bool includePrivateParameters) {
            if (includePrivateParameters) { PrivateExportCalls++; throw new CryptographicException("Private parameters cannot be exported."); }
            return _inner.ExportParameters(false);
        }
        public override void ImportParameters(RSAParameters parameters) => throw new NotSupportedException();
        public override byte[] SignHash(byte[] hash, HashAlgorithmName hashAlgorithm, RSASignaturePadding padding) => _inner.SignHash(hash, hashAlgorithm, padding);
        public override bool VerifyHash(byte[] hash, byte[] signature, HashAlgorithmName hashAlgorithm, RSASignaturePadding padding) => _inner.VerifyHash(hash, signature, hashAlgorithm, padding);
    }
}
