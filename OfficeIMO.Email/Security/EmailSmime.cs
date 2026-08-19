using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.Email;

/// <summary>Thin S/MIME verification and decryption orchestration over the shared OfficeIMO security engine.</summary>
public static partial class EmailSmime {
    /// <summary>Verifies clear-signed or opaque-signed S/MIME content retained by the email reader.</summary>
    public static EmailSmimeVerificationResult Verify(
        EmailDocument document,
        IOfficeSecurityProvider securityProvider,
        CmsVerificationOptions? options = null,
        EmailReaderOptions? contentReaderOptions = null) {
#if NETSTANDARD2_0 || NET472
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
#else
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(securityProvider);
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
                securityProvider.Name,
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
            return new EmailSmimeVerificationResult(securityProvider.Name,
                document.Protection.Kind, null, null, null, diagnostics);
        }

        CmsVerificationResult cryptography = payload!.DetachedContent == null
            ? securityProvider.VerifyCms(
                payload.EncodedCms,
                options,
                CertificateValidationPurpose.EmailSigning)
            : VerifyDetached(securityProvider, payload, options, diagnostics);
        AddTrustDiagnostics(cryptography, options, diagnostics);
        byte[]? signedMimeEntity = payload.DetachedContent ?? cryptography.EncapsulatedContent;
        EmailDocument? signedContent = TryParseContent(
            signedMimeEntity,
            contentReaderOptions,
            diagnostics,
            "message/protection/signed-content");
        return new EmailSmimeVerificationResult(
            securityProvider.Name,
            document.Protection.Kind,
            cryptography,
            signedMimeEntity,
            signedContent,
            diagnostics);
    }

    private static CmsVerificationResult VerifyDetached(
        IOfficeSecurityProvider securityProvider,
        MimeSmimeExtractor.ExtractedSmimePayload payload,
        CmsVerificationOptions options,
        List<EmailDiagnostic> diagnostics) {
        byte[] original = payload.DetachedContent!;
        CmsVerificationResult exact = securityProvider.VerifyCmsDetached(
            payload.EncodedCms,
            original,
            options,
            CertificateValidationPurpose.EmailSigning);
        if (exact.IsCryptographicallyValid ||
            !MimeSmimeExtractor.TryCanonicalizeLineEndings(
                original,
                options.MaxContentBytes,
                out byte[] canonical)) {
            return exact;
        }

        CmsVerificationResult normalized = securityProvider.VerifyCmsDetached(
            payload.EncodedCms,
            canonical,
            options,
            CertificateValidationPurpose.EmailSigning);
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
        IOfficeSecurityProvider securityProvider,
        CmsEnvelopeOptions? options = null,
        EmailReaderOptions? contentReaderOptions = null) {
#if NETSTANDARD2_0 || NET472
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (recipientCertificate == null) throw new ArgumentNullException(nameof(recipientCertificate));
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
#else
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(recipientCertificate);
        ArgumentNullException.ThrowIfNull(securityProvider);
#endif
        options ??= new CmsEnvelopeOptions();
        var diagnostics = new List<EmailDiagnostic>();
        if (document.Protection.Kind != EmailProtectionKind.SmimeOpaque) {
            diagnostics.Add(new EmailDiagnostic(
                "EMAIL_SMIME_ENVELOPE_NOT_DETECTED",
                "The document is not classified as opaque S/MIME content.",
                EmailDiagnosticSeverity.Warning,
                "message/protection"));
            return new EmailSmimeDecryptionResult(securityProvider.Name,
                document.Protection.Kind, null, null, null, diagnostics);
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
            return new EmailSmimeDecryptionResult(securityProvider.Name,
                document.Protection.Kind, null, null, null, diagnostics);
        }

        CmsDecryptionResult cryptography = securityProvider.DecryptCms(
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
            securityProvider.Name,
            document.Protection.Kind,
            cryptography,
            decrypted,
            decryptedContent,
            diagnostics);
    }

    /// <summary>
    /// Decrypts outer EnvelopedData and only then verifies an inner clear- or opaque-signed MIME entity. The concrete
    /// provider is supplied explicitly and the original protected source remains available on the input document.
    /// </summary>
    public static EmailSmimeProcessingResult DecryptThenVerify(
        EmailDocument document,
        X509Certificate2 recipientCertificate,
        IOfficeSecurityProvider securityProvider,
        CmsEnvelopeOptions? envelopeOptions = null,
        CmsVerificationOptions? verificationOptions = null,
        EmailReaderOptions? contentReaderOptions = null) {
#if NETSTANDARD2_0 || NET472
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (recipientCertificate == null) throw new ArgumentNullException(nameof(recipientCertificate));
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
#else
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(recipientCertificate);
        ArgumentNullException.ThrowIfNull(securityProvider);
#endif
        EmailSmimeDecryptionResult decryption = Decrypt(document, recipientCertificate,
            securityProvider, envelopeOptions, contentReaderOptions);
        var diagnostics = new List<EmailDiagnostic>(decryption.Diagnostics);
        var order = new List<EmailSmimeProcessingStage> { EmailSmimeProcessingStage.Decrypt };
        EmailSmimeVerificationResult? verification = null;
        EmailDocument? content = decryption.DecryptedContent;
        CmsVerificationOptions effectiveVerificationOptions =
            verificationOptions ?? new CmsVerificationOptions();
        OpaqueCmsContentKind opaqueContentKind = content?.Protection.Kind == EmailProtectionKind.SmimeOpaque
            ? MimeSmimeExtractor.ClassifyOpaqueCms(content,
                effectiveVerificationOptions.MaxEncodedBytes,
                effectiveVerificationOptions.MaxContentBytes,
                diagnostics)
            : OpaqueCmsContentKind.Unknown;
        if (content != null &&
            (content.Protection.Kind == EmailProtectionKind.SmimeClearSigned ||
             opaqueContentKind == OpaqueCmsContentKind.SignedData)) {
            diagnostics.Add(new EmailDiagnostic(
                EmailSmimeDiagnosticCodes.DecryptThenVerify,
                "The outer S/MIME envelope was decrypted before the inner protected MIME entity was verified.",
                EmailDiagnosticSeverity.Information,
                "message/protection"));
            verification = Verify(content, securityProvider, effectiveVerificationOptions, contentReaderOptions);
            order.Add(EmailSmimeProcessingStage.Verify);
            diagnostics.AddRange(verification.Diagnostics);
            content = verification.SignedContent ?? content;
        } else if (content?.Protection.Kind == EmailProtectionKind.SmimeOpaque) {
            string message = opaqueContentKind == OpaqueCmsContentKind.EnvelopedData
                ? "The decrypted inner opaque S/MIME entity contains enveloped-data. It was retained without adding a verification stage."
                : "The decrypted inner opaque S/MIME entity could not be classified as signed-data. It was retained without adding a verification stage.";
            diagnostics.Add(new EmailDiagnostic(
                EmailSmimeDiagnosticCodes.InnerOpaqueNotSigned,
                message,
                EmailDiagnosticSeverity.Information,
                "message/protection/decrypted-content"));
        }
        return new EmailSmimeProcessingResult(securityProvider.Name, decryption, verification,
            content, order.AsReadOnly(), diagnostics.AsReadOnly());
    }

    private static void AddTrustDiagnostics(CmsVerificationResult result,
        CmsVerificationOptions options, ICollection<EmailDiagnostic> diagnostics) {
        if (options.CertificateValidation.DisableCertificateDownloads) {
            diagnostics.Add(new EmailDiagnostic(
                EmailSmimeDiagnosticCodes.OfflinePolicy,
                "Certificate downloads were disabled. Chain, revocation, and timestamp results reflect the supplied offline trust material and revocation policy.",
                EmailDiagnosticSeverity.Information,
                "message/protection/trust"));
        }
        foreach (CmsSignerVerificationResult signer in result.Signers) {
            string location = "message/protection/signers/" + signer.SignerIndex.ToString(CultureInfo.InvariantCulture);
            diagnostics.Add(new EmailDiagnostic(
                EmailSmimeDiagnosticCodes.SignerIdentity,
                string.Concat("Signer subject: ", signer.Subject ?? "unknown", "; issuer: ",
                    signer.Issuer ?? "unknown", "; serial: ", signer.SerialNumber ?? "unknown", "."),
                EmailDiagnosticSeverity.Information,
                location));
            AddStatusDiagnostic(diagnostics, EmailSmimeDiagnosticCodes.ChainStatus,
                "Certificate chain", signer.CertificateValidation.ChainStatus, location + "/chain");
            AddStatusDiagnostic(diagnostics, EmailSmimeDiagnosticCodes.RevocationStatus,
                "Certificate revocation", signer.CertificateValidation.RevocationStatus, location + "/revocation");
            AddStatusDiagnostic(diagnostics, EmailSmimeDiagnosticCodes.TimestampStatus,
                "Signature timestamp", signer.TimestampStatus, location + "/timestamp");
        }
    }

    private static void AddStatusDiagnostic(ICollection<EmailDiagnostic> diagnostics, string code,
        string label, SecurityValidationStatus status, string location) {
        EmailDiagnosticSeverity severity = status == SecurityValidationStatus.Invalid
            ? EmailDiagnosticSeverity.Error
            : status == SecurityValidationStatus.Indeterminate
                ? EmailDiagnosticSeverity.Warning
                : EmailDiagnosticSeverity.Information;
        diagnostics.Add(new EmailDiagnostic(code,
            string.Concat(label, " status: ", status.ToString(), "."), severity, location));
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
