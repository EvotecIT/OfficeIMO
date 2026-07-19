using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Security;

namespace OfficeIMO.Security;

/// <summary>Creates and decrypts CMS EnvelopedData for X.509 recipients.</summary>
public static class CmsEnvelopedDataService {
    /// <summary>Encrypts content for every supplied recipient certificate.</summary>
    public static byte[] Encrypt(
        byte[] content,
        IEnumerable<X509Certificate2> recipients,
        CmsEnvelopeOptions? options = null) {
#if NETSTANDARD2_0 || NET472
        if (content == null) throw new ArgumentNullException(nameof(content));
        if (recipients == null) throw new ArgumentNullException(nameof(recipients));
#else
        ArgumentNullException.ThrowIfNull(content);
        ArgumentNullException.ThrowIfNull(recipients);
#endif
        options ??= new CmsEnvelopeOptions();
        SecurityLimits.EnsureBufferWithinLimit(content, options.MaxContentBytes, nameof(content));
        List<X509Certificate2> recipientList = recipients.ToList();
        if (recipientList.Count == 0) throw new ArgumentException("At least one recipient certificate is required.", nameof(recipients));
        SecurityLimits.EnsureCountWithinLimit(recipientList.Count, options.MaxRecipients, nameof(options.MaxRecipients));

        var generator = new CmsEnvelopedDataGenerator();
        foreach (X509Certificate2 certificate in recipientList) {
            generator.AddKeyTransRecipient(DotNetUtilities.FromX509Certificate(certificate));
        }

        var processable = new CmsProcessableByteArray(content);
        string algorithm = GetContentEncryptionAlgorithm(options.ContentEncryptionAlgorithm);
        return generator.Generate(processable, algorithm).GetEncoded();
    }

    /// <summary>Decrypts content for a matching RSA recipient certificate.</summary>
    public static CmsDecryptionResult Decrypt(
        byte[] encodedCms,
        X509Certificate2 recipientCertificate,
        CmsEnvelopeOptions? options = null) {
#if NETSTANDARD2_0 || NET472
        if (encodedCms == null) throw new ArgumentNullException(nameof(encodedCms));
        if (recipientCertificate == null) throw new ArgumentNullException(nameof(recipientCertificate));
#else
        ArgumentNullException.ThrowIfNull(encodedCms);
        ArgumentNullException.ThrowIfNull(recipientCertificate);
#endif
        options ??= new CmsEnvelopeOptions();
        SecurityLimits.EnsureBufferWithinLimit(encodedCms, options.MaxEncodedBytes, nameof(encodedCms));
        var findings = new List<SecurityFinding>();
        try {
            var envelope = new CmsEnvelopedData(encodedCms);
            IList<RecipientInformation> recipients = envelope.GetRecipientInfos().GetRecipients();
            SecurityLimits.EnsureCountWithinLimit(recipients.Count, options.MaxRecipients, nameof(options.MaxRecipients));
            Org.BouncyCastle.X509.X509Certificate bcCertificate =
                DotNetUtilities.FromX509Certificate(recipientCertificate);
            RecipientInformation? recipient = recipients.FirstOrDefault(item => item.RecipientID.Match(bcCertificate));
            if (recipient == null) {
                findings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Error,
                    "EnvelopeRecipientNotFound",
                    "The supplied certificate does not match any CMS recipient identifier."));
                return Failure(parsed: true, envelope.EncryptionAlgorithmID.Algorithm.Id, null, findings);
            }

            using RSA? rsa = recipientCertificate.GetRSAPrivateKey();
            if (rsa == null) {
                findings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Error,
                    "EnvelopePrivateKeyMissing",
                    "The matching recipient certificate does not expose an RSA private key."));
                return Failure(
                    parsed: true,
                    envelope.EncryptionAlgorithmID.Algorithm.Id,
                    recipient.KeyEncryptionAlgorithmID.Algorithm.Id,
                    findings);
            }

            Org.BouncyCastle.Crypto.AsymmetricKeyParameter privateKey;
            try {
                privateKey = DotNetUtilities.GetRsaKeyPair(rsa).Private;
            } catch (Exception exception) when (exception is CryptographicException or NotSupportedException) {
                findings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Error,
                    "EnvelopePrivateKeyNotExportable",
                    "CMS decryption requires an exportable RSA private key with the current Bouncy Castle recipient adapter: " +
                    exception.Message));
                return Failure(
                    parsed: true,
                    envelope.EncryptionAlgorithmID.Algorithm.Id,
                    recipient.KeyEncryptionAlgorithmID.Algorithm.Id,
                    findings);
            }

            byte[] decrypted = recipient.GetContent(privateKey);
            SecurityLimits.EnsureBufferWithinLimit(decrypted, options.MaxContentBytes, nameof(options.MaxContentBytes));
            return new CmsDecryptionResult(
                parsed: true,
                decrypted: true,
                decrypted,
                envelope.EncryptionAlgorithmID.Algorithm.Id,
                recipient.KeyEncryptionAlgorithmID.Algorithm.Id,
                findings);
        } catch (Exception exception) when (IsValidationException(exception)) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                "EnvelopeMalformedOrUndecryptable",
                "The CMS envelope could not be decoded or decrypted: " + exception.Message));
            return Failure(parsed: false, null, null, findings);
        }
    }

    private static string GetContentEncryptionAlgorithm(CmsContentEncryptionAlgorithm algorithm) => algorithm switch {
        CmsContentEncryptionAlgorithm.Aes128Cbc => CmsEnvelopedGenerator.Aes128Cbc,
        CmsContentEncryptionAlgorithm.Aes192Cbc => CmsEnvelopedGenerator.Aes192Cbc,
        CmsContentEncryptionAlgorithm.Aes256Cbc => CmsEnvelopedGenerator.Aes256Cbc,
        _ => throw new ArgumentOutOfRangeException(nameof(algorithm), algorithm, "Unsupported CMS content-encryption algorithm.")
    };

    private static CmsDecryptionResult Failure(
        bool parsed,
        string? contentEncryptionAlgorithmOid,
        string? keyEncryptionAlgorithmOid,
        IReadOnlyList<SecurityFinding> findings) =>
        new CmsDecryptionResult(
            parsed,
            decrypted: false,
            content: null,
            contentEncryptionAlgorithmOid,
            keyEncryptionAlgorithmOid,
            findings);

    private static bool IsValidationException(Exception exception) =>
        exception is not OutOfMemoryException &&
        exception is not StackOverflowException &&
        exception is not AccessViolationException;
}
