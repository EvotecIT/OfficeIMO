using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.Email;

public static partial class EmailSmime {
    /// <summary>Creates a clear-signed or opaque-signed S/MIME message.</summary>
    public static EmailSmimeCreationResult Sign(
        EmailDocument document,
        X509Certificate2 signingCertificate,
        IOfficeSecurityProvider securityProvider,
        EmailSmimeSignatureMode mode = EmailSmimeSignatureMode.ClearSigned,
        CmsSigningOptions? options = null,
        IEnumerable<X509Certificate2>? certificateChain = null,
        EmailWriterOptions? writerOptions = null) {
        ValidateCreationArguments(document, securityProvider);
        if (signingCertificate == null) throw new ArgumentNullException(nameof(signingCertificate));
        options ??= new CmsSigningOptions();
        EmailWriterOptions writer = writerOptions ?? EmailWriterOptions.Default;
        var diagnostics = new List<EmailDiagnostic>();
        byte[] content = MimeWriter.WriteMimeEntity(document, writer, diagnostics);
        return CreateSignedResult(document, signingCertificate, securityProvider, mode, options,
            certificateChain, writer, diagnostics, content);
    }

    /// <summary>Creates an encrypted S/MIME EnvelopedData message for the supplied recipients.</summary>
    public static EmailSmimeCreationResult Encrypt(
        EmailDocument document,
        IEnumerable<X509Certificate2> recipients,
        IOfficeSecurityProvider securityProvider,
        CmsEnvelopeOptions? options = null,
        EmailWriterOptions? writerOptions = null) {
        ValidateCreationArguments(document, securityProvider);
        if (recipients == null) throw new ArgumentNullException(nameof(recipients));
        options ??= new CmsEnvelopeOptions();
        X509Certificate2[] recipientArray = MaterializeRecipients(recipients, options.MaxRecipients);
        EmailWriterOptions writer = writerOptions ?? EmailWriterOptions.Default;
        var diagnostics = new List<EmailDiagnostic>();
        byte[] content = MimeWriter.WriteMimeEntity(document, writer, diagnostics);
        byte[] cms = securityProvider.EncryptCms(content, recipientArray, options);
        byte[] protectedEntity = CreateOpaqueEntity(cms, "enveloped-data", writer.Base64LineLength,
            writer.MaxOutputBytes);
        byte[] message = ComposeMessage(document, writer, protectedEntity);
        return new EmailSmimeCreationResult(EmailProtectionKind.SmimeOpaque, message, protectedEntity, cms,
            diagnostics.AsReadOnly());
    }

    /// <summary>Signs the MIME entity and then encrypts the complete signed entity for the supplied recipients.</summary>
    public static EmailSmimeCreationResult SignAndEncrypt(
        EmailDocument document,
        X509Certificate2 signingCertificate,
        IEnumerable<X509Certificate2> recipients,
        IOfficeSecurityProvider securityProvider,
        EmailSmimeSignatureMode signatureMode = EmailSmimeSignatureMode.ClearSigned,
        CmsSigningOptions? signingOptions = null,
        CmsEnvelopeOptions? envelopeOptions = null,
        IEnumerable<X509Certificate2>? certificateChain = null,
        EmailWriterOptions? writerOptions = null) {
        ValidateCreationArguments(document, securityProvider);
        if (signingCertificate == null) throw new ArgumentNullException(nameof(signingCertificate));
        if (recipients == null) throw new ArgumentNullException(nameof(recipients));
        signingOptions ??= new CmsSigningOptions();
        envelopeOptions ??= new CmsEnvelopeOptions();
        X509Certificate2[] recipientArray = MaterializeRecipients(recipients, envelopeOptions.MaxRecipients);
        EmailWriterOptions writer = writerOptions ?? EmailWriterOptions.Default;
        var diagnostics = new List<EmailDiagnostic>();
        byte[] content = MimeWriter.WriteMimeEntity(document, writer, diagnostics);
        byte[] signedEntity = CreateSignedEntity(signingCertificate, securityProvider, signatureMode, signingOptions,
            certificateChain, writer, content, out _, out _);
        byte[] cms = securityProvider.EncryptCms(signedEntity, recipientArray, envelopeOptions);
        byte[] protectedEntity = CreateOpaqueEntity(cms, "enveloped-data", writer.Base64LineLength,
            writer.MaxOutputBytes);
        byte[] message = ComposeMessage(document, writer, protectedEntity);
        return new EmailSmimeCreationResult(EmailProtectionKind.SmimeOpaque, message, protectedEntity, cms,
            diagnostics.AsReadOnly());
    }

    private static EmailSmimeCreationResult CreateSignedResult(EmailDocument document,
        X509Certificate2 signingCertificate, IOfficeSecurityProvider securityProvider, EmailSmimeSignatureMode mode,
        CmsSigningOptions options, IEnumerable<X509Certificate2>? certificateChain, EmailWriterOptions writer,
        List<EmailDiagnostic> diagnostics, byte[] content) {
        byte[] protectedEntity = CreateSignedEntity(signingCertificate, securityProvider, mode, options,
            certificateChain, writer, content, out byte[] cms, out EmailProtectionKind kind);
        byte[] message = ComposeMessage(document, writer, protectedEntity);
        return new EmailSmimeCreationResult(kind, message, protectedEntity, cms, diagnostics.AsReadOnly());
    }

    private static X509Certificate2[] MaterializeRecipients(IEnumerable<X509Certificate2> recipients,
        int maxRecipients) {
        if (maxRecipients <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxRecipients), maxRecipients,
                "The S/MIME recipient limit must be positive.");
        }
        var materialized = new List<X509Certificate2>(Math.Min(maxRecipients, 16));
        foreach (X509Certificate2 certificate in recipients) {
            if (certificate == null) {
                throw new ArgumentException("S/MIME recipient certificates cannot contain null entries.", nameof(recipients));
            }
            materialized.Add(certificate);
            if (materialized.Count > maxRecipients) {
                throw new EmailLimitExceededException(nameof(CmsEnvelopeOptions.MaxRecipients),
                    materialized.Count, maxRecipients);
            }
        }
        if (materialized.Count == 0) {
            throw new ArgumentException("At least one S/MIME recipient certificate is required.", nameof(recipients));
        }
        return materialized.ToArray();
    }

    private static byte[] CreateSignedEntity(X509Certificate2 signingCertificate,
        IOfficeSecurityProvider securityProvider, EmailSmimeSignatureMode mode, CmsSigningOptions options,
        IEnumerable<X509Certificate2>? certificateChain, EmailWriterOptions writer, byte[] content,
        out byte[] cms, out EmailProtectionKind kind) {
        byte[] protectedEntity;
        switch (mode) {
            case EmailSmimeSignatureMode.ClearSigned:
                byte[] detachedContent = RemoveTrailingCrlf(content);
                cms = securityProvider.SignCmsDetached(detachedContent, signingCertificate, options, certificateChain);
                protectedEntity = CreateClearSignedEntity(detachedContent, cms, options.DigestAlgorithm,
                    writer.Base64LineLength, writer.MaxOutputBytes);
                kind = EmailProtectionKind.SmimeClearSigned;
                break;
            case EmailSmimeSignatureMode.OpaqueSigned:
                cms = securityProvider.SignCmsEncapsulated(content, signingCertificate, options, certificateChain);
                protectedEntity = CreateOpaqueEntity(cms, "signed-data", writer.Base64LineLength,
                    writer.MaxOutputBytes);
                kind = EmailProtectionKind.SmimeOpaque;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(mode), mode, "Unknown S/MIME signature mode.");
        }
        return protectedEntity;
    }

    private static byte[] CreateClearSignedEntity(byte[] content, byte[] signature, HashAlgorithmName digestAlgorithm,
        int lineLength, long maxOutputBytes) {
        string boundary = "officeimo-smime-" + Guid.NewGuid().ToString("N");
        using var output = new EmailBoundedMemoryStream(maxOutputBytes);
        WriteAscii(output, "Content-Type: multipart/signed; protocol=\"application/pkcs7-signature\"; micalg=" +
            GetMicalg(digestAlgorithm) + "; boundary=\"" + boundary + "\"\r\n\r\n--" + boundary + "\r\n");
        output.Write(content, 0, content.Length);
        EnsureEndsWithCrlf(output);
        WriteAscii(output,
            "--" + boundary + "\r\n" +
            "Content-Type: application/pkcs7-signature; name=smime.p7s\r\n" +
            "Content-Transfer-Encoding: base64\r\n" +
            "Content-Disposition: attachment; filename=smime.p7s\r\n\r\n");
        WriteFoldedBase64(output, signature, lineLength);
        WriteAscii(output, "--" + boundary + "--\r\n");
        return output.ToArray();
    }

    private static byte[] CreateOpaqueEntity(byte[] cms, string smimeType, int lineLength, long maxOutputBytes) {
        using var output = new EmailBoundedMemoryStream(maxOutputBytes);
        WriteAscii(output,
            "Content-Type: application/pkcs7-mime; smime-type=" + smimeType + "; name=smime.p7m\r\n" +
            "Content-Transfer-Encoding: base64\r\n" +
            "Content-Disposition: attachment; filename=smime.p7m\r\n\r\n");
        WriteFoldedBase64(output, cms, lineLength);
        return output.ToArray();
    }

    private static byte[] ComposeMessage(EmailDocument document, EmailWriterOptions writer, byte[] protectedEntity) {
        byte[] headers = MimeWriter.WriteTransportHeaders(document, writer);
        using var output = new EmailBoundedMemoryStream(writer.MaxOutputBytes);
        output.Write(headers, 0, headers.Length);
        output.Write(protectedEntity, 0, protectedEntity.Length);
        return output.ToArray();
    }

    private static void WriteFoldedBase64(Stream output, byte[] value, int lineLength) {
        int inputBytesPerLine = lineLength / 4 * 3;
        for (int offset = 0; offset < value.Length; offset += inputBytesPerLine) {
            int count = Math.Min(inputBytesPerLine, value.Length - offset);
            WriteAscii(output, Convert.ToBase64String(value, offset, count) + "\r\n");
        }
    }

    private static string GetMicalg(HashAlgorithmName algorithm) {
        if (algorithm == HashAlgorithmName.SHA256) return "sha-256";
        if (algorithm == HashAlgorithmName.SHA384) return "sha-384";
        if (algorithm == HashAlgorithmName.SHA512) return "sha-512";
        if (algorithm == HashAlgorithmName.SHA1) return "sha-1";
        throw new NotSupportedException($"S/MIME micalg does not support digest algorithm '{algorithm.Name}'.");
    }

    private static void EnsureEndsWithCrlf(MemoryStream output) {
        byte[] value = output.GetBuffer();
        long length = output.Length;
        if (length >= 2 && value[length - 2] == (byte)'\r' && value[length - 1] == (byte)'\n') return;
        WriteAscii(output, "\r\n");
    }

    private static byte[] RemoveTrailingCrlf(byte[] value) {
        if (value.Length < 2 || value[value.Length - 2] != (byte)'\r' || value[value.Length - 1] != (byte)'\n') {
            return value;
        }
        var result = new byte[value.Length - 2];
        Buffer.BlockCopy(value, 0, result, 0, result.Length);
        return result;
    }

    private static void WriteAscii(Stream output, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        output.Write(bytes, 0, bytes.Length);
    }

    private static void ValidateCreationArguments(EmailDocument document, IOfficeSecurityProvider securityProvider) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
    }
}
