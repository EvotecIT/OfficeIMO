using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.Email;

/// <summary>Thin S/MIME verification and decryption orchestration over the shared OfficeIMO security engine.</summary>
public static class EmailSmime {
    /// <summary>Verifies clear-signed or opaque-signed S/MIME content retained by the email reader.</summary>
    public static EmailSmimeVerificationResult Verify(
        EmailDocument document,
        CmsVerificationOptions? options = null,
        EmailReaderOptions? contentReaderOptions = null) {
#if NETSTANDARD2_0 || NET472
        if (document == null) throw new ArgumentNullException(nameof(document));
#else
        ArgumentNullException.ThrowIfNull(document);
#endif
        options ??= new CmsVerificationOptions();
        var diagnostics = new List<EmailDiagnostic>();
        if (document.Protection.Kind != EmailProtectionKind.SmimeClearSigned &&
            document.Protection.Kind != EmailProtectionKind.SmimeOpaque) {
            diagnostics.Add(new EmailDiagnostic(
                "EMAIL_SMIME_NOT_DETECTED",
                "The document is not classified as clear-signed or opaque S/MIME content.",
                EmailDiagnosticSeverity.Warning,
                "message/protection"));
            return new EmailSmimeVerificationResult(
                document.Protection.Kind,
                null,
                null,
                null,
                diagnostics);
        }

        if (!MimeSmimeExtractor.TryExtract(
                document,
                options.MaxEncodedBytes,
                options.MaxContentBytes,
                diagnostics,
                out MimeSmimeExtractor.ExtractedSmimePayload? payload)) {
            diagnostics.Add(new EmailDiagnostic(
                "EMAIL_SMIME_PAYLOAD_MISSING",
                "No retained CMS payload was available for S/MIME verification.",
                EmailDiagnosticSeverity.Error,
                "message/protection"));
            return new EmailSmimeVerificationResult(document.Protection.Kind, null, null, null, diagnostics);
        }

        CmsVerificationResult cryptography = payload!.DetachedContent == null
            ? CmsSignedDataVerifier.Verify(payload.EncodedCms, options)
            : VerifyDetached(payload, options, diagnostics);
        byte[]? signedMimeEntity = payload.DetachedContent ?? cryptography.EncapsulatedContent;
        EmailDocument? signedContent = TryParseContent(
            signedMimeEntity,
            contentReaderOptions,
            diagnostics,
            "message/protection/signed-content");
        return new EmailSmimeVerificationResult(
            document.Protection.Kind,
            cryptography,
            signedMimeEntity,
            signedContent,
            diagnostics);
    }

    private static CmsVerificationResult VerifyDetached(
        MimeSmimeExtractor.ExtractedSmimePayload payload,
        CmsVerificationOptions options,
        List<EmailDiagnostic> diagnostics) {
        byte[] original = payload.DetachedContent!;
        CmsVerificationResult exact = CmsSignedDataVerifier.VerifyDetached(payload.EncodedCms, original, options);
        if (exact.IsCryptographicallyValid ||
            !MimeSmimeExtractor.TryCanonicalizeLineEndings(
                original,
                options.MaxContentBytes,
                out byte[] canonical)) {
            return exact;
        }

        CmsVerificationResult normalized = CmsSignedDataVerifier.VerifyDetached(
            payload.EncodedCms,
            canonical,
            options);
        if (!normalized.IsCryptographicallyValid) return exact;
        diagnostics.Add(new EmailDiagnostic(
            "EMAIL_SMIME_CANONICAL_LINE_ENDINGS_APPLIED",
            "The detached MIME entity used non-canonical line endings. Its signature validated after standard MIME CRLF canonicalization; SignedMimeEntity retains the original source bytes.",
            EmailDiagnosticSeverity.Information,
            "message/protection/signed-content"));
        return normalized;
    }

    /// <summary>Decrypts opaque S/MIME EnvelopedData for a matching recipient certificate.</summary>
    public static EmailSmimeDecryptionResult Decrypt(
        EmailDocument document,
        X509Certificate2 recipientCertificate,
        CmsEnvelopeOptions? options = null,
        EmailReaderOptions? contentReaderOptions = null) {
#if NETSTANDARD2_0 || NET472
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (recipientCertificate == null) throw new ArgumentNullException(nameof(recipientCertificate));
#else
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(recipientCertificate);
#endif
        options ??= new CmsEnvelopeOptions();
        var diagnostics = new List<EmailDiagnostic>();
        if (document.Protection.Kind != EmailProtectionKind.SmimeOpaque) {
            diagnostics.Add(new EmailDiagnostic(
                "EMAIL_SMIME_ENVELOPE_NOT_DETECTED",
                "The document is not classified as opaque S/MIME content.",
                EmailDiagnosticSeverity.Warning,
                "message/protection"));
            return new EmailSmimeDecryptionResult(document.Protection.Kind, null, null, null, diagnostics);
        }

        if (!MimeSmimeExtractor.TryExtract(
                document,
                options.MaxEncodedBytes,
                options.MaxContentBytes,
                diagnostics,
                out MimeSmimeExtractor.ExtractedSmimePayload? payload)) {
            diagnostics.Add(new EmailDiagnostic(
                "EMAIL_SMIME_PAYLOAD_MISSING",
                "No retained CMS payload was available for S/MIME decryption.",
                EmailDiagnosticSeverity.Error,
                "message/protection"));
            return new EmailSmimeDecryptionResult(document.Protection.Kind, null, null, null, diagnostics);
        }

        CmsDecryptionResult cryptography = CmsEnvelopedDataService.Decrypt(
            payload!.EncodedCms,
            recipientCertificate,
            options);
        byte[]? decrypted = cryptography.Content;
        EmailDocument? decryptedContent = TryParseContent(
            decrypted,
            contentReaderOptions,
            diagnostics,
            "message/protection/decrypted-content");
        return new EmailSmimeDecryptionResult(
            document.Protection.Kind,
            cryptography,
            decrypted,
            decryptedContent,
            diagnostics);
    }

    private static EmailDocument? TryParseContent(
        byte[]? content,
        EmailReaderOptions? readerOptions,
        List<EmailDiagnostic> diagnostics,
        string location) {
        if (content == null || content.Length == 0) return null;
        try {
            using EmailReadResult read = new EmailDocumentReader(readerOptions ?? EmailReaderOptions.Default).Read(content);
            foreach (EmailDiagnostic diagnostic in read.Diagnostics) diagnostics.Add(diagnostic);
            return read.Document;
        } catch (Exception exception) when (exception is InvalidDataException or EmailLimitExceededException or NotSupportedException) {
            diagnostics.Add(new EmailDiagnostic(
                "EMAIL_SMIME_CONTENT_PARSE_FAILED",
                "The protected MIME content was retained but could not be projected: " + exception.Message,
                EmailDiagnosticSeverity.Warning,
                location));
            return null;
        }
    }
}
