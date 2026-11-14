using OfficeIMO.Security;

namespace OfficeIMO.Email;

/// <summary>Inline ASCII-armored OpenPGP content found at the start of a plain-text body.</summary>
public enum EmailInlinePgpKind {
    /// <summary>No inline OpenPGP block was detected.</summary>
    None = 0,
    /// <summary>A cleartext-signed message ("BEGIN PGP SIGNED MESSAGE").</summary>
    CleartextSigned = 1,
    /// <summary>An armored OpenPGP message ("BEGIN PGP MESSAGE"), encrypted or signed.</summary>
    Message = 2
}

/// <summary>
/// PGP/MIME (RFC 3156) and inline OpenPGP verification and decryption over a caller-supplied OpenPGP provider and key
/// set. Nothing reads a user keyring; only the supplied keys are used.
/// </summary>
public static class EmailPgp {
    /// <summary>
    /// Verifies a PGP/MIME multipart/signed message over the exact bytes of its signed part, or an inline
    /// cleartext-signed body.
    /// </summary>
    public static EmailPgpVerificationResult Verify(
        EmailDocument document,
        IOfficeOpenPgpProvider provider,
        OfficeOpenPgpKeySet keys,
        OfficeOpenPgpOptions? options = null,
        EmailReaderOptions? contentReaderOptions = null) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(provider);
        ArgumentNullException.ThrowIfNull(keys);
        options ??= new OfficeOpenPgpOptions();
        var diagnostics = new List<EmailDiagnostic>();
        EmailProtectionKind kind = document.Protection.Kind;
        EmailInlinePgpKind inline = document.Protection.InlinePgp;

        if (kind == EmailProtectionKind.PgpMimeClearSigned) {
            byte[]? entity;
            byte[]? signature;
            string? micalg;
            try {
                if (!MimePgpExtractor.TryExtractSigned(document, options.MaxInputBytes, diagnostics, out entity, out signature, out micalg)) {
                    return new EmailPgpVerificationResult(provider.Name, kind, inline, null, null, null, null, diagnostics);
                }
            } catch (Exception exception) when (exception is InvalidDataException or IOException) {
                diagnostics.Add(Error("EMAIL_PGP_PAYLOAD_UNAVAILABLE", "The PGP/MIME payload could not be read: " + exception.Message));
                return new EmailPgpVerificationResult(provider.Name, kind, inline, null, null, null, null, diagnostics);
            }

            OfficeOpenPgpVerificationResult result = provider.VerifyDetached(signature!, entity!, keys, options);
            if (!result.IsValid && MimeSmimeExtractor.TryCanonicalizeLineEndings(entity!, options.MaxInputBytes, out byte[] canonical)) {
                OfficeOpenPgpVerificationResult retried = provider.VerifyDetached(signature!, canonical, keys, options);
                if (retried.IsValid) {
                    result = retried;
                    diagnostics.Add(new EmailDiagnostic("EMAIL_PGP_CANONICAL_LINE_ENDINGS_APPLIED",
                        "The signed MIME part used bare LF line endings. Its signature verified after RFC 3156 CRLF canonicalization; SignedMimeEntity keeps the source bytes.",
                        EmailDiagnosticSeverity.Information, "message/protection/signed-content"));
                }
            }

            CheckMicalg(micalg, result, diagnostics);
            EmailDocument? content = TryParse(entity, contentReaderOptions, diagnostics, "message/protection/signed-content");
            return new EmailPgpVerificationResult(provider.Name, kind, inline, result, entity, content, null, diagnostics);
        }

        if (inline == EmailInlinePgpKind.CleartextSigned) {
            diagnostics.Add(InlineSourceNote());
            OfficeOpenPgpVerificationResult result = provider.VerifyCleartext(Encoding.UTF8.GetBytes(document.Body.Text ?? string.Empty), keys, options);
            string? text = result.SignedContent == null ? null : Encoding.UTF8.GetString(result.SignedContent);
            return new EmailPgpVerificationResult(provider.Name, kind, inline, result, null, null, text, diagnostics);
        }

        diagnostics.Add(new EmailDiagnostic("EMAIL_PGP_SIGNATURE_NOT_DETECTED",
            "The document is neither a PGP/MIME multipart/signed message nor an inline cleartext-signed body.",
            EmailDiagnosticSeverity.Warning, "message/protection"));
        return new EmailPgpVerificationResult(provider.Name, kind, inline, null, null, null, null, diagnostics);
    }

    /// <summary>
    /// Decrypts a PGP/MIME multipart/encrypted message or an inline armored message with a supplied secret key. A
    /// decrypted PGP/MIME multipart/signed entity is verified as well (RFC 3156 section 6.1).
    /// </summary>
    public static EmailPgpDecryptionResult Decrypt(
        EmailDocument document,
        IOfficeOpenPgpProvider provider,
        OfficeOpenPgpKeySet keys,
        OfficeOpenPgpOptions? options = null,
        EmailReaderOptions? contentReaderOptions = null) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(provider);
        ArgumentNullException.ThrowIfNull(keys);
        options ??= new OfficeOpenPgpOptions();
        var diagnostics = new List<EmailDiagnostic>();
        EmailProtectionKind kind = document.Protection.Kind;
        EmailInlinePgpKind inline = document.Protection.InlinePgp;

        if (kind == EmailProtectionKind.PgpMimeEncrypted) {
            byte[]? control;
            byte[]? payload;
            try {
                if (!MimePgpExtractor.TryExtractEncrypted(document, options.MaxInputBytes, diagnostics, out control, out payload)) {
                    return new EmailPgpDecryptionResult(provider.Name, kind, inline, null, null, null, null, null, diagnostics);
                }
            } catch (Exception exception) when (exception is InvalidDataException or IOException) {
                diagnostics.Add(Error("EMAIL_PGP_PAYLOAD_UNAVAILABLE", "The PGP/MIME payload could not be read: " + exception.Message));
                return new EmailPgpDecryptionResult(provider.Name, kind, inline, null, null, null, null, null, diagnostics);
            }

            if (!MimePgpExtractor.HasVersionOne(control!)) {
                diagnostics.Add(Error("EMAIL_PGP_CONTROL_INVALID",
                    "The application/pgp-encrypted control part does not contain \"Version: 1\" (RFC 3156 section 4)."));
                return new EmailPgpDecryptionResult(provider.Name, kind, inline, null, null, null, null, null, diagnostics);
            }

            OfficeOpenPgpMessageResult result = provider.OpenMessage(payload!, keys, options);
            if (result.Status != OfficeOpenPgpMessageStatus.Opened) {
                return new EmailPgpDecryptionResult(provider.Name, kind, inline, result, null, null, null, null, diagnostics);
            }

            EmailReaderOptions readerOptions = contentReaderOptions ?? new EmailReaderOptions(preserveRawSource: true);
            EmailDocument? content = TryParse(result.Content, readerOptions, diagnostics, "message/protection/decrypted-content");
            EmailPgpVerificationResult? nested = content?.Protection.Kind == EmailProtectionKind.PgpMimeClearSigned
                ? Verify(content, provider, keys, options, contentReaderOptions)
                : null;
            return new EmailPgpDecryptionResult(provider.Name, kind, inline, result, result.Content, content, null, nested, diagnostics);
        }

        if (inline == EmailInlinePgpKind.Message) {
            diagnostics.Add(InlineSourceNote());
            OfficeOpenPgpMessageResult result = provider.OpenMessage(Encoding.UTF8.GetBytes(document.Body.Text ?? string.Empty), keys, options);
            string? text = result.Content == null ? null : Encoding.UTF8.GetString(result.Content);
            return new EmailPgpDecryptionResult(provider.Name, kind, inline, result, null, null, text, null, diagnostics);
        }

        diagnostics.Add(new EmailDiagnostic("EMAIL_PGP_ENCRYPTION_NOT_DETECTED",
            "The document is neither a PGP/MIME multipart/encrypted message nor an inline armored OpenPGP message.",
            EmailDiagnosticSeverity.Warning, "message/protection"));
        return new EmailPgpDecryptionResult(provider.Name, kind, inline, null, null, null, null, null, diagnostics);
    }

    private static void CheckMicalg(string? micalg, OfficeOpenPgpVerificationResult result, List<EmailDiagnostic> diagnostics) {
        if (micalg == null || result.Signatures.Count == 0) return;
        foreach (OfficeOpenPgpSignatureInfo signature in result.Signatures) {
            if (signature.HashAlgorithm == null) continue;
            string expected = "pgp-" + signature.HashAlgorithm.ToLowerInvariant();
            if (!micalg.Split(',').Any(value => value.Trim().Equals(expected, StringComparison.OrdinalIgnoreCase))) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_PGP_MICALG_MISMATCH",
                    "The micalg parameter '" + micalg + "' does not name the signature hash " + expected + ".",
                    EmailDiagnosticSeverity.Warning, "message/protection"));
            }
        }
    }

    private static EmailDiagnostic InlineSourceNote() => new EmailDiagnostic("EMAIL_PGP_INLINE_TEXT",
        "Inline OpenPGP was processed from the decoded body text encoded as UTF-8. Signatures made over another charset will not verify.",
        EmailDiagnosticSeverity.Information, "message/body/text");

    private static EmailDiagnostic Error(string code, string message) =>
        new EmailDiagnostic(code, message, EmailDiagnosticSeverity.Error, "message/protection");

    private static EmailDocument? TryParse(byte[]? content, EmailReaderOptions? readerOptions, List<EmailDiagnostic> diagnostics,
        string location) {
        if (content == null || content.Length == 0) return null;
        try {
            using EmailReadResult read = new EmailDocumentReader(readerOptions ?? EmailReaderOptions.Default).ReadDocument(content);
            foreach (EmailDiagnostic diagnostic in read.Diagnostics) diagnostics.Add(diagnostic);
            return read.Document;
        } catch (Exception exception) when (exception is InvalidDataException or EmailLimitExceededException or NotSupportedException) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_PGP_CONTENT_PARSE_FAILED",
                "The protected MIME content was retained but could not be projected: " + exception.Message,
                EmailDiagnosticSeverity.Warning, location));
            return null;
        }
    }
}

/// <summary>Result of verifying a PGP/MIME or inline cleartext OpenPGP signature.</summary>
public sealed class EmailPgpVerificationResult {
    internal EmailPgpVerificationResult(string providerName, EmailProtectionKind protectionKind, EmailInlinePgpKind inlinePgp,
        OfficeOpenPgpVerificationResult? cryptography, byte[]? signedMimeEntity, EmailDocument? signedContent, string? signedText,
        IReadOnlyList<EmailDiagnostic> diagnostics) {
        ProviderName = providerName;
        ProtectionKind = protectionKind;
        InlinePgp = inlinePgp;
        Cryptography = cryptography;
        SignedMimeEntity = signedMimeEntity;
        SignedContent = signedContent;
        SignedText = signedText;
        Diagnostics = diagnostics;
    }

    /// <summary>OpenPGP provider that produced the evidence.</summary>
    public string ProviderName { get; }

    /// <summary>MIME protection wrapper detected by the reader.</summary>
    public EmailProtectionKind ProtectionKind { get; }

    /// <summary>Inline OpenPGP block detected in the body.</summary>
    public EmailInlinePgpKind InlinePgp { get; }

    /// <summary>Signature evaluation, or null when nothing verifiable was found.</summary>
    public OfficeOpenPgpVerificationResult? Cryptography { get; }

    /// <summary>Exact bytes of the signed MIME part, as extracted from the source.</summary>
    public byte[]? SignedMimeEntity { get; }

    /// <summary>Parsed signed MIME part.</summary>
    public EmailDocument? SignedContent { get; }

    /// <summary>For inline cleartext signatures, exactly the signed text.</summary>
    public string? SignedText { get; }

    /// <summary>Extraction and projection diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>True when at least one signature was found and all signatures are valid.</summary>
    public bool IsValid => Cryptography?.IsValid == true;
}

/// <summary>Result of decrypting a PGP/MIME or inline OpenPGP message.</summary>
public sealed class EmailPgpDecryptionResult {
    internal EmailPgpDecryptionResult(string providerName, EmailProtectionKind protectionKind, EmailInlinePgpKind inlinePgp,
        OfficeOpenPgpMessageResult? cryptography, byte[]? decryptedMimeEntity, EmailDocument? decryptedContent, string? decryptedText,
        EmailPgpVerificationResult? nestedSignature, IReadOnlyList<EmailDiagnostic> diagnostics) {
        ProviderName = providerName;
        ProtectionKind = protectionKind;
        InlinePgp = inlinePgp;
        Cryptography = cryptography;
        DecryptedMimeEntity = decryptedMimeEntity;
        DecryptedContent = decryptedContent;
        DecryptedText = decryptedText;
        NestedSignature = nestedSignature;
        Diagnostics = diagnostics;
    }

    /// <summary>OpenPGP provider that produced the evidence.</summary>
    public string ProviderName { get; }

    /// <summary>MIME protection wrapper detected by the reader.</summary>
    public EmailProtectionKind ProtectionKind { get; }

    /// <summary>Inline OpenPGP block detected in the body.</summary>
    public EmailInlinePgpKind InlinePgp { get; }

    /// <summary>Decryption result, including signatures combined with the encryption (RFC 3156 section 6.2).</summary>
    public OfficeOpenPgpMessageResult? Cryptography { get; }

    /// <summary>Decrypted MIME entity bytes for PGP/MIME messages.</summary>
    public byte[]? DecryptedMimeEntity { get; }

    /// <summary>Parsed decrypted MIME entity.</summary>
    public EmailDocument? DecryptedContent { get; }

    /// <summary>Decrypted text for inline messages.</summary>
    public string? DecryptedText { get; }

    /// <summary>Verification of a multipart/signed entity found inside the encryption (RFC 3156 section 6.1).</summary>
    public EmailPgpVerificationResult? NestedSignature { get; }

    /// <summary>Extraction and projection diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>True when the content was decrypted with integrity protection (or read, for signed-only inline messages).</summary>
    public bool IsDecrypted => Cryptography?.Status == OfficeOpenPgpMessageStatus.Opened;

    /// <summary>True when a combined or nested signature exists and every such signature is valid.</summary>
    public bool IsSignatureValid {
        get {
            bool combined = Cryptography?.Signatures.Count > 0;
            bool nested = NestedSignature?.Cryptography?.Signatures.Count > 0;
            if (!combined && !nested) return false;
            return (!combined || Cryptography!.IsSignatureValid) && (!nested || NestedSignature!.IsValid);
        }
    }
}
