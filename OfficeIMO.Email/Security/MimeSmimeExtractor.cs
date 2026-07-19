namespace OfficeIMO.Email;

internal static class MimeSmimeExtractor {
    internal static bool TryExtract(
        EmailDocument document,
        long maximumCmsBytes,
        List<EmailDiagnostic> diagnostics,
        out ExtractedSmimePayload? payload) {
        payload = null;
        if (document.RawSource != null && document.Format == EmailFileFormat.Eml &&
            TryExtractFromMime(document.RawSource, diagnostics, out payload)) {
            return true;
        }

        EmailAttachment? attachment = document.Protection.PayloadAttachment;
        if (attachment == null) return false;
        try {
            byte[] encoded = ReadAttachment(attachment, maximumCmsBytes);
            if (document.Protection.Kind == EmailProtectionKind.SmimeClearSigned ||
                (attachment.ContentType ?? string.Empty)
                    .IndexOf("multipart/signed", StringComparison.OrdinalIgnoreCase) >= 0) {
                if (TryExtractFromMime(encoded, diagnostics, out payload)) return true;
                diagnostics.Add(new EmailDiagnostic(
                    "EMAIL_SMIME_SIGNED_ENTITY_INVALID",
                    "The retained Outlook S/MIME attachment is not a complete multipart/signed MIME entity.",
                    EmailDiagnosticSeverity.Error,
                    "message/protection/payload"));
                return false;
            }
            payload = new ExtractedSmimePayload(encoded, null);
            return true;
        } catch (Exception exception) when (exception is IOException or InvalidDataException or UnauthorizedAccessException) {
            diagnostics.Add(new EmailDiagnostic(
                "EMAIL_SMIME_PAYLOAD_UNAVAILABLE",
                "The projected S/MIME payload could not be read: " + exception.Message,
                EmailDiagnosticSeverity.Error,
                "message/protection"));
            return false;
        }
    }

    private static bool TryExtractFromMime(
        byte[] source,
        List<EmailDiagnostic> diagnostics,
        out ExtractedSmimePayload? payload) {
        payload = null;
        var headers = new List<EmailHeader>();
        EmailReaderOptions options = EmailReaderOptions.Default;
        int bodyOffset = MimeHeaderParser.Parse(
            source,
            0,
            source.Length,
            options,
            headers,
            diagnostics,
            "message");
        MimeValue contentType = MimeValueParser.Parse(
            MimeHeaderParser.GetValue(headers, "Content-Type"),
            "text/plain",
            diagnostics,
            "message");
        string? transferEncoding = MimeHeaderParser.GetValue(headers, "Content-Transfer-Encoding");
        int bodyCount = source.Length - bodyOffset;

        if (string.Equals(contentType.Value, "multipart/signed", StringComparison.OrdinalIgnoreCase) &&
            (contentType.GetParameter("protocol") ?? string.Empty)
                .IndexOf("pkcs7-signature", StringComparison.OrdinalIgnoreCase) >= 0) {
            string? boundary = contentType.GetParameter("boundary");
            if (boundary == null) return false;
            var state = new MimeParserState(options, diagnostics, CancellationToken.None);
            List<ArraySegment<byte>> parts = MimeParser.SplitMultipart(
                source,
                bodyOffset,
                bodyCount,
                boundary,
                state,
                "message");
            if (parts.Count < 2) {
                diagnostics.Add(new EmailDiagnostic(
                    "EMAIL_SMIME_MULTIPART_INCOMPLETE",
                    "The S/MIME multipart/signed entity does not contain both content and signature parts.",
                    EmailDiagnosticSeverity.Error,
                    "message"));
                return false;
            }

            byte[] signedEntity = Copy(parts[0]);
            byte[] signature = DecodePart(parts[1], source, options, diagnostics, "message/signature");
            payload = new ExtractedSmimePayload(signature, signedEntity);
            return true;
        }

        if (string.Equals(contentType.Value, "application/pkcs7-mime", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(contentType.Value, "application/x-pkcs7-mime", StringComparison.OrdinalIgnoreCase)) {
            byte[] body = Copy(source, bodyOffset, bodyCount);
            byte[] encoded = MimeTextCodec.DecodeTransfer(body, transferEncoding, diagnostics, "message");
            payload = new ExtractedSmimePayload(encoded, null);
            return true;
        }
        return false;
    }

    private static byte[] DecodePart(
        ArraySegment<byte> part,
        byte[] source,
        EmailReaderOptions options,
        IList<EmailDiagnostic> diagnostics,
        string location) {
        var headers = new List<EmailHeader>();
        int bodyOffset = MimeHeaderParser.Parse(
            source,
            part.Offset,
            part.Count,
            options,
            headers,
            diagnostics,
            location);
        int end = part.Offset + part.Count;
        byte[] body = Copy(source, bodyOffset, Math.Max(0, end - bodyOffset));
        return MimeTextCodec.DecodeTransfer(
            body,
            MimeHeaderParser.GetValue(headers, "Content-Transfer-Encoding"),
            diagnostics,
            location);
    }

    private static byte[] ReadAttachment(EmailAttachment attachment, long maximumBytes) {
        if (attachment.Content != null) {
            if (attachment.Content.LongLength > maximumBytes) {
                throw new InvalidDataException("The S/MIME payload exceeds the configured CMS limit.");
            }
            return (byte[])attachment.Content.Clone();
        }
        using Stream stream = attachment.OpenContentStream();
        return EmailByteReader.ReadAll(stream, maximumBytes, CancellationToken.None);
    }

    private static byte[] Copy(ArraySegment<byte> value) => Copy(value.Array!, value.Offset, value.Count);

    private static byte[] Copy(byte[] source, int offset, int count) {
        var result = new byte[count];
        Buffer.BlockCopy(source, offset, result, 0, count);
        return result;
    }

    internal sealed class ExtractedSmimePayload {
        internal ExtractedSmimePayload(byte[] encodedCms, byte[]? detachedContent) {
            EncodedCms = encodedCms;
            DetachedContent = detachedContent;
        }

        internal byte[] EncodedCms { get; }
        internal byte[]? DetachedContent { get; }
    }
}
