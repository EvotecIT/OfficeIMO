using OfficeIMO.Security;

namespace OfficeIMO.Email;

/// <summary>
/// DKIM (RFC 6376, RFC 8463, RFC 8301) verification over the exact received message bytes. Public keys come from
/// a caller-supplied resolver and signatures are checked by a caller-supplied verifier such as
/// <c>OfficeIMO.Security.OfficeSecurityProvider</c>.
/// </summary>
public static class EmailDkim {
    /// <summary>Verifies every DKIM-Signature header field in a raw RFC 5322 message.</summary>
    public static async Task<EmailDkimVerificationResult> VerifyAsync(
        byte[] rawMessage,
        IEmailDkimKeyResolver keyResolver,
        IOfficePublicKeySignatureVerifier signatureVerifier,
        EmailAuthenticationOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(rawMessage);
        ArgumentNullException.ThrowIfNull(keyResolver);
        ArgumentNullException.ThrowIfNull(signatureVerifier);
        options ??= new EmailAuthenticationOptions();
        options.Validate();
        var diagnostics = new List<EmailDiagnostic>();
        EmailAuthenticationMessage? message = EmailAuthenticationParsing.Parse(rawMessage, options, diagnostics, "dkim");
        var signatures = new List<EmailDkimSignatureResult>();
        if (message == null) return new EmailDkimVerificationResult(signatures, diagnostics);

        var evaluator = new EmailDkimEvaluator(message, keyResolver, signatureVerifier, options);
        for (int index = 0; index < message.Fields.Count; index++) {
            if (!message.Fields[index].Is("DKIM-Signature")) continue;
            if (signatures.Count >= options.MaxSignatures) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_DKIM_SIGNATURE_LIMIT",
                    "More DKIM-Signature header fields are present than the configured limit; later ones were not evaluated.",
                    EmailDiagnosticSeverity.Warning, "headers/dkim-signature"));
                break;
            }

            signatures.Add(await evaluator.EvaluateAsync(index, EmailDkimSignatureKind.DkimSignature, null, cancellationToken)
                .ConfigureAwait(false));
        }

        return new EmailDkimVerificationResult(signatures, diagnostics);
    }

    /// <summary>
    /// Verifies DKIM signatures over the source bytes retained by the reader. Read the message with
    /// <see cref="EmailReaderOptions.PreserveRawSource"/> enabled; a regenerated or modified model cannot be verified.
    /// </summary>
    public static Task<EmailDkimVerificationResult> VerifyAsync(
        EmailDocument document,
        IEmailDkimKeyResolver keyResolver,
        IOfficePublicKeySignatureVerifier signatureVerifier,
        EmailAuthenticationOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        if (document.RawSource == null) {
            return Task.FromResult(new EmailDkimVerificationResult(Array.Empty<EmailDkimSignatureResult>(),
                new[] { EmailAuthenticationParsing.RawSourceMissing("dkim") }));
        }

        return VerifyAsync(document.RawSource, keyResolver, signatureVerifier, options, cancellationToken);
    }
}

internal static class EmailAuthenticationParsing {
    internal static EmailAuthenticationMessage? Parse(byte[] raw, EmailAuthenticationOptions options,
        List<EmailDiagnostic> diagnostics, string area) {
        if (raw.LongLength > options.MaxMessageBytes) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_AUTH_MESSAGE_TOO_LARGE",
                "The message exceeds the configured size limit and was not verified.",
                EmailDiagnosticSeverity.Error, area));
            return null;
        }

        EmailAuthenticationMessage message = EmailAuthenticationMessage.Parse(raw, options.MaxHeaderFields, out bool headerLimitExceeded);
        if (headerLimitExceeded) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_AUTH_HEADER_LIMIT",
                "The message has more header fields than the configured limit and was not verified.",
                EmailDiagnosticSeverity.Error, area));
            return null;
        }

        if (message.LineEndingsNormalized) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_AUTH_LINE_ENDINGS_NORMALIZED",
                "The message uses bare LF line endings; they were treated as CRLF as on the SMTP wire.",
                EmailDiagnosticSeverity.Information, area));
        }

        return message;
    }

    internal static EmailDiagnostic RawSourceMissing(string area) => new EmailDiagnostic("EMAIL_AUTH_RAW_SOURCE_MISSING",
        "The document does not retain its original bytes. Read it with PreserveRawSource enabled to verify transport signatures.",
        EmailDiagnosticSeverity.Error, area);
}
