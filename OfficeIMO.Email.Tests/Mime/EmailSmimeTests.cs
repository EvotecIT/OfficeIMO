#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Email;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailSmimeTests {
    [Fact]
    public void ClearSignedMessage_VerifiesTheExactFirstMimeEntity() {
        using X509Certificate2 certificate = CreateCertificate("OfficeIMO S-MIME Signer");
        byte[] signedEntity = Encoding.ASCII.GetBytes(
            "Content-Type: text/plain; charset=utf-8\r\n" +
            "Content-Transfer-Encoding: 7bit\r\n\r\n" +
            "Hello from clear-signed S/MIME");
        byte[] signature = CmsSignedDataSigner.SignDetached(signedEntity, certificate);
        byte[] message = CreateClearSignedMessage(signedEntity, signature);
        using EmailReadResult read = new EmailDocumentReader().Read(message);

        EmailSmimeVerificationResult result = EmailSmime.Verify(read.Document, TrustSelfSigned());

        Assert.Equal(EmailProtectionKind.SmimeClearSigned, result.ProtectionKind);
        Assert.True(result.IsCryptographicallyValid);
        Assert.Equal(signedEntity, result.SignedMimeEntity);
        Assert.Equal("Hello from clear-signed S/MIME", result.SignedContent?.Body.Text);
    }

    [Fact]
    public void ClearSignedMessage_DetectsContentTampering() {
        using X509Certificate2 certificate = CreateCertificate("OfficeIMO S-MIME Signer");
        byte[] signedEntity = Encoding.ASCII.GetBytes("Content-Type: text/plain\r\n\r\nOriginal");
        byte[] signature = CmsSignedDataSigner.SignDetached(signedEntity, certificate);
        byte[] tamperedEntity = Encoding.ASCII.GetBytes("Content-Type: text/plain\r\n\r\nTampered");
        using EmailReadResult read = new EmailDocumentReader().Read(
            CreateClearSignedMessage(tamperedEntity, signature));

        EmailSmimeVerificationResult result = EmailSmime.Verify(read.Document, TrustSelfSigned());

        Assert.False(result.IsCryptographicallyValid);
        Assert.Equal(SecurityValidationStatus.Invalid, result.Cryptography?.Signers[0].DigestStatus);
    }

    [Fact]
    public void ClearSignedMessage_VerifiesCanonicalMimeAfterTransportLineEndingNormalization() {
        using X509Certificate2 certificate = CreateCertificate("OfficeIMO canonical S-MIME signer");
        byte[] canonical = Encoding.ASCII.GetBytes(
            "Content-Type: text/plain; charset=utf-8\r\n" +
            "Content-Transfer-Encoding: 7bit\r\n\r\n" +
            "Outlook canonical content\r\n");
        byte[] signature = CmsSignedDataSigner.SignDetached(canonical, certificate);
        byte[] normalizedByTransport = Encoding.ASCII.GetBytes(
            Encoding.ASCII.GetString(CreateClearSignedMessage(canonical, signature))
                .Replace("\r\n", "\n"));
        using EmailReadResult read = new EmailDocumentReader().Read(normalizedByTransport);

        EmailSmimeVerificationResult result = EmailSmime.Verify(read.Document, TrustSelfSigned());

        Assert.True(result.IsCryptographicallyValid);
        Assert.NotNull(result.SignedMimeEntity);
        Assert.DoesNotContain((byte)'\r', result.SignedMimeEntity!);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_SMIME_CANONICAL_LINE_ENDINGS_APPLIED");
        Assert.Equal("Outlook canonical content\n", result.SignedContent?.Body.Text);
    }

    [Fact]
    public void ClearSignedMessage_DoesNotCanonicalizeBeyondConfiguredContentLimit() {
        using X509Certificate2 certificate = CreateCertificate("OfficeIMO bounded canonical S-MIME signer");
        byte[] canonical = Encoding.ASCII.GetBytes(
            "Content-Type: text/plain; charset=utf-8\r\n" +
            "Content-Transfer-Encoding: 7bit\r\n\r\n" +
            "Bounded canonical content\r\n");
        byte[] signature = CmsSignedDataSigner.SignDetached(canonical, certificate);
        byte[] normalizedEntity = Encoding.ASCII.GetBytes(
            Encoding.ASCII.GetString(canonical).Replace("\r\n", "\n"));
        byte[] normalizedByTransport = Encoding.ASCII.GetBytes(
            Encoding.ASCII.GetString(CreateClearSignedMessage(canonical, signature))
                .Replace("\r\n", "\n"));
        using EmailReadResult read = new EmailDocumentReader().Read(normalizedByTransport);
        CmsVerificationOptions options = TrustSelfSigned();
        options.MaxContentBytes = normalizedEntity.LongLength;

        EmailSmimeVerificationResult result = EmailSmime.Verify(read.Document, options);

        Assert.False(result.IsCryptographicallyValid);
        Assert.DoesNotContain(result.Diagnostics, diagnostic =>
            diagnostic.Code == "EMAIL_SMIME_CANONICAL_LINE_ENDINGS_APPLIED");
    }

    [Theory]
    [InlineData(EmailFileFormat.OutlookMsg)]
    [InlineData(EmailFileFormat.Tnef)]
    public void OutlookClearSignedAttachment_VerifiesTheRetainedMultipartEntity(EmailFileFormat format) {
        using X509Certificate2 certificate = CreateCertificate("OfficeIMO Outlook S-MIME Signer");
        byte[] signedEntity = Encoding.ASCII.GetBytes(
            "Content-Type: text/plain; charset=utf-8\r\n\r\n" +
            "Outlook retained clear-signed body");
        byte[] signature = CmsSignedDataSigner.SignDetached(signedEntity, certificate);
        byte[] protectedEntity = CreateClearSignedMimeEntity(signedEntity, signature);
        var source = new EmailDocument {
            Format = format,
            MessageClass = "IPM.Note.SMIME.MultipartSigned",
            Subject = "Outlook clear signed"
        };
        source.Attachments.Add(new EmailAttachment {
            FileName = "SMIME.p7m",
            ContentType = "multipart/signed",
            Content = protectedEntity,
            Length = protectedEntity.Length,
            MapiAttachMethod = 1
        });
        using EmailReadResult read = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, format));

        EmailSmimeVerificationResult result = EmailSmime.Verify(read.Document, TrustSelfSigned());

        Assert.Equal(EmailProtectionKind.SmimeClearSigned, result.ProtectionKind);
        Assert.True(result.IsCryptographicallyValid);
        Assert.Equal(signedEntity, result.SignedMimeEntity);
        Assert.Equal("Outlook retained clear-signed body", result.SignedContent?.Body.Text);
    }

    [Fact]
    public void OpaqueSignedMessage_VerifiesAndProjectsEncapsulatedMimeContent() {
        using X509Certificate2 certificate = CreateCertificate("OfficeIMO Opaque Signer");
        byte[] content = Encoding.ASCII.GetBytes("Content-Type: text/plain\r\n\r\nOpaque signed body");
        byte[] cms = CmsSignedDataSigner.SignEncapsulated(content, certificate);
        using EmailReadResult read = new EmailDocumentReader().Read(CreateOpaqueMessage(cms, "signed-data"));

        EmailSmimeVerificationResult result = EmailSmime.Verify(read.Document, TrustSelfSigned());

        Assert.Equal(EmailProtectionKind.SmimeOpaque, result.ProtectionKind);
        Assert.True(result.IsCryptographicallyValid);
        Assert.Equal(content, result.SignedMimeEntity);
        Assert.Equal("Opaque signed body", result.SignedContent?.Body.Text);
    }

    [Fact]
    public void EnvelopedMessage_DecryptsAndProjectsMimeContent() {
        using X509Certificate2 recipient = CreateCertificate("OfficeIMO S-MIME Recipient");
        byte[] content = Encoding.ASCII.GetBytes(
            "Content-Type: text/plain; charset=utf-8\r\n\r\nConfidential message body");
        byte[] cms = CmsEnvelopedDataService.Encrypt(content, new[] { recipient });
        using EmailReadResult read = new EmailDocumentReader().Read(CreateOpaqueMessage(cms, "enveloped-data"));

        EmailSmimeDecryptionResult result = EmailSmime.Decrypt(read.Document, recipient);

        Assert.True(result.Decrypted);
        Assert.Equal(content, result.DecryptedMimeEntity);
        Assert.Equal("Confidential message body", result.DecryptedContent?.Body.Text);
    }

    [Fact]
    public void EnvelopedMessage_ReportsWrongRecipientWithoutLosingThePayload() {
        using X509Certificate2 recipient = CreateCertificate("OfficeIMO S-MIME Recipient");
        using X509Certificate2 other = CreateCertificate("OfficeIMO Other Recipient");
        byte[] cms = CmsEnvelopedDataService.Encrypt(
            Encoding.ASCII.GetBytes("Content-Type: text/plain\r\n\r\nSecret"),
            new[] { recipient });
        using EmailReadResult read = new EmailDocumentReader().Read(CreateOpaqueMessage(cms, "enveloped-data"));

        EmailSmimeDecryptionResult result = EmailSmime.Decrypt(read.Document, other);

        Assert.False(result.Decrypted);
        Assert.Contains(result.Cryptography!.Findings, finding => finding.Code == "EnvelopeRecipientNotFound");
        Assert.NotNull(read.Document.RawSource);
    }

    private static byte[] CreateClearSignedMessage(byte[] signedEntity, byte[] signature) {
        byte[] prefix = Encoding.ASCII.GetBytes(
            "From: sender@example.test\r\n" +
            "To: recipient@example.test\r\n" +
            "Subject: Clear signed\r\n" +
            "MIME-Version: 1.0\r\n");
        return Combine(prefix, CreateClearSignedMimeEntity(signedEntity, signature));
    }

    private static byte[] CreateClearSignedMimeEntity(byte[] signedEntity, byte[] signature) {
        byte[] prefix = Encoding.ASCII.GetBytes(
            "Content-Type: multipart/signed; protocol=\"application/pkcs7-signature\"; micalg=sha-256; boundary=\"officeimo-sig\"\r\n\r\n" +
            "--officeimo-sig\r\n");
        byte[] separator = Encoding.ASCII.GetBytes(
            "\r\n--officeimo-sig\r\n" +
            "Content-Type: application/pkcs7-signature; name=smime.p7s\r\n" +
            "Content-Transfer-Encoding: base64\r\n" +
            "Content-Disposition: attachment; filename=smime.p7s\r\n\r\n" +
            Convert.ToBase64String(signature) + "\r\n" +
            "--officeimo-sig--\r\n");
        return Combine(prefix, signedEntity, separator);
    }

    private static byte[] CreateOpaqueMessage(byte[] cms, string smimeType) => Encoding.ASCII.GetBytes(
        "From: sender@example.test\r\n" +
        "To: recipient@example.test\r\n" +
        "Subject: Opaque S-MIME\r\n" +
        "MIME-Version: 1.0\r\n" +
        "Content-Type: application/pkcs7-mime; smime-type=" + smimeType + "; name=smime.p7m\r\n" +
        "Content-Transfer-Encoding: base64\r\n" +
        "Content-Disposition: attachment; filename=smime.p7m\r\n\r\n" +
        Convert.ToBase64String(cms) + "\r\n");

    private static byte[] Combine(params byte[][] values) {
        var result = new byte[values.Sum(static value => value.Length)];
        int offset = 0;
        foreach (byte[] value in values) {
            Buffer.BlockCopy(value, 0, result, offset, value.Length);
            offset += value.Length;
        }
        return result;
    }

    private static CmsVerificationOptions TrustSelfSigned() {
        var options = new CmsVerificationOptions();
        options.CertificateValidation.ChainEvaluator = static (_, _) => true;
        return options;
    }

    private static X509Certificate2 CreateCertificate(string commonName) {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=" + commonName,
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(
            X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.KeyEncipherment,
            critical: true));
        request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
            new OidCollection {
                new Oid("1.3.6.1.5.5.7.3.4")
            },
            critical: false));
        request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, critical: false));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }
}
#endif
