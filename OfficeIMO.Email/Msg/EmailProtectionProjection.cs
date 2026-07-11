namespace OfficeIMO.Email;

internal static class EmailProtectionProjection {
    internal static void Apply(EmailDocument document, IList<EmailDiagnostic> diagnostics, string location) {
        string messageClass = document.MessageClass ?? string.Empty;
        EmailProtectionKind kind = messageClass.IndexOf("SMIME.MultipartSigned", StringComparison.OrdinalIgnoreCase) >= 0
            ? EmailProtectionKind.SmimeClearSigned
            : messageClass.IndexOf("SMIME", StringComparison.OrdinalIgnoreCase) >= 0
                ? EmailProtectionKind.SmimeOpaque
                : EmailProtectionKind.None;
        document.Protection.Kind = kind;
        document.Protection.MessageClass = document.MessageClass;
        if (kind == EmailProtectionKind.None) return;

        EmailAttachment? payload = document.Attachments.FirstOrDefault(IsProtectedPayload);
        document.Protection.PayloadAttachment = payload;
        if (payload == null) {
            diagnostics.Add(new EmailDiagnostic(
                "EMAIL_PROTECTED_PAYLOAD_MISSING",
                "The Outlook message class indicates S/MIME protection, but no CMS or signed MIME payload attachment was found.",
                EmailDiagnosticSeverity.Warning,
                location));
        }
    }

    private static bool IsProtectedPayload(EmailAttachment attachment) {
        string fileName = attachment.FileName ?? string.Empty;
        string contentType = attachment.ContentType ?? string.Empty;
        return fileName.EndsWith(".p7m", StringComparison.OrdinalIgnoreCase) ||
            fileName.EndsWith(".p7s", StringComparison.OrdinalIgnoreCase) ||
            contentType.IndexOf("pkcs7", StringComparison.OrdinalIgnoreCase) >= 0 ||
            contentType.IndexOf("multipart/signed", StringComparison.OrdinalIgnoreCase) >= 0;
    }
}
