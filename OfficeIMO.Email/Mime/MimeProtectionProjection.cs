namespace OfficeIMO.Email;

internal static class MimeProtectionProjection {
    internal static void Apply(EmailDocument document, IReadOnlyList<EmailHeader> headers,
        IList<EmailDiagnostic> diagnostics, string location) {
        MimeValue contentType = MimeValueParser.Parse(MimeHeaderParser.GetValue(headers, "Content-Type"),
            "text/plain", diagnostics, location);
        string protocol = contentType.GetParameter("protocol") ?? string.Empty;
        EmailProtectionKind kind = Classify(contentType.Value, protocol);
        if (kind == EmailProtectionKind.None) return;

        document.Protection.Kind = kind;
        document.Protection.MessageClass = document.MessageClass;
        document.Protection.PayloadAttachment = kind == EmailProtectionKind.PgpMimeEncrypted ||
            kind == EmailProtectionKind.MimeEncrypted
            ? document.Attachments.LastOrDefault(attachment => !IsEncryptionControlPart(attachment))
            : document.Attachments.FirstOrDefault(IsProtectedPayload);
        if (document.Protection.PayloadAttachment == null &&
            !contentType.Value.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase)) {
            document.Protection.PayloadAttachment = document.Attachments.FirstOrDefault();
        }
    }

    private static EmailProtectionKind Classify(string contentType, string protocol) {
        if (string.Equals(contentType, "application/pkcs7-mime", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(contentType, "application/x-pkcs7-mime", StringComparison.OrdinalIgnoreCase)) {
            return EmailProtectionKind.SmimeOpaque;
        }
        if (string.Equals(contentType, "multipart/signed", StringComparison.OrdinalIgnoreCase)) {
            if (protocol.IndexOf("pkcs7-signature", StringComparison.OrdinalIgnoreCase) >= 0) {
                return EmailProtectionKind.SmimeClearSigned;
            }
            if (protocol.IndexOf("pgp-signature", StringComparison.OrdinalIgnoreCase) >= 0) {
                return EmailProtectionKind.PgpMimeClearSigned;
            }
            return EmailProtectionKind.MimeClearSigned;
        }
        if (string.Equals(contentType, "multipart/encrypted", StringComparison.OrdinalIgnoreCase)) {
            return protocol.IndexOf("pgp-encrypted", StringComparison.OrdinalIgnoreCase) >= 0
                ? EmailProtectionKind.PgpMimeEncrypted
                : EmailProtectionKind.MimeEncrypted;
        }
        return EmailProtectionKind.None;
    }

    private static bool IsProtectedPayload(EmailAttachment attachment) {
        string contentType = attachment.ContentType ?? string.Empty;
        string fileName = attachment.FileName ?? string.Empty;
        return contentType.IndexOf("pkcs7", StringComparison.OrdinalIgnoreCase) >= 0 ||
            contentType.IndexOf("pgp-signature", StringComparison.OrdinalIgnoreCase) >= 0 ||
            contentType.IndexOf("pgp-encrypted", StringComparison.OrdinalIgnoreCase) >= 0 ||
            fileName.EndsWith(".p7m", StringComparison.OrdinalIgnoreCase) ||
            fileName.EndsWith(".p7s", StringComparison.OrdinalIgnoreCase) ||
            fileName.EndsWith(".asc", StringComparison.OrdinalIgnoreCase) ||
            fileName.EndsWith(".pgp", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsEncryptionControlPart(EmailAttachment attachment) =>
        string.Equals(attachment.ContentType, "application/pgp-encrypted", StringComparison.OrdinalIgnoreCase);
}
