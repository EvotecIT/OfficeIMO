namespace OfficeIMO.Email;

internal static class MimeSmimeExtractor {
    internal static bool TryExtract(
        EmailDocument document,
        long maximumCmsBytes,
        long maximumContentBytes,
        List<EmailDiagnostic> diagnostics,
        out ExtractedSmimePayload? payload) {
        payload = null;
        if (document.RawSource != null && document.Format == EmailFileFormat.Eml) {
            try {
                if (TryExtractFromMime(document.RawSource, maximumCmsBytes, maximumContentBytes,
                        diagnostics, out payload)) {
                    return true;
                }
            } catch (Exception exception) when (exception is IOException or InvalidDataException or UnauthorizedAccessException) {
                diagnostics.Add(new EmailDiagnostic(
                    "EMAIL_SMIME_PAYLOAD_UNAVAILABLE",
                    "The retained S/MIME payload could not be read: " + exception.Message,
                    EmailDiagnosticSeverity.Error,
                    "message/protection"));
                return false;
            }
        }

        EmailAttachment? attachment = document.Protection.PayloadAttachment;
        if (attachment == null) return false;
        try {
            byte[] encoded = ReadAttachment(attachment, maximumCmsBytes);
            if (document.Protection.Kind == EmailProtectionKind.SmimeClearSigned ||
                (attachment.ContentType ?? string.Empty)
                    .IndexOf("multipart/signed", StringComparison.OrdinalIgnoreCase) >= 0) {
                if (TryExtractFromMime(encoded, maximumCmsBytes, maximumContentBytes,
                        diagnostics, out payload)) return true;
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
        long maximumCmsBytes,
        long maximumContentBytes,
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

            if (parts[0].Count > maximumContentBytes) {
                throw new InvalidDataException(
                    "The S/MIME signed content exceeds the configured CMS content limit.");
            }
            byte[] signedEntity = Copy(parts[0]);
            byte[] signature = DecodePart(parts[1], source, options, diagnostics,
                "message/signature", maximumCmsBytes);
            payload = new ExtractedSmimePayload(signature, signedEntity);
            return true;
        }

        if (string.Equals(contentType.Value, "application/pkcs7-mime", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(contentType.Value, "application/x-pkcs7-mime", StringComparison.OrdinalIgnoreCase)) {
            byte[] encoded = DecodeTransferPart(source, bodyOffset, bodyCount,
                transferEncoding, diagnostics, "message", maximumCmsBytes);
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
        string location,
        long maximumBytes) {
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
        return DecodeTransferPart(
            source,
            bodyOffset,
            Math.Max(0, end - bodyOffset),
            MimeHeaderParser.GetValue(headers, "Content-Transfer-Encoding"),
            diagnostics,
            location,
            maximumBytes);
    }

    private static byte[] DecodeTransferPart(
        byte[] source,
        int offset,
        int count,
        string? transferEncoding,
        IList<EmailDiagnostic> diagnostics,
        string location,
        long maximumBytes) {
        var preflightDiagnostics = new List<EmailDiagnostic>();
        long decodedLength = MimeTextCodec.GetDecodedLength(
            source, offset, count, transferEncoding, preflightDiagnostics, location);
        if (decodedLength > maximumBytes) {
            foreach (EmailDiagnostic diagnostic in preflightDiagnostics) diagnostics.Add(diagnostic);
            throw new InvalidDataException("The S/MIME payload exceeds the configured CMS limit.");
        }
        byte[] body = Copy(source, offset, count);
        byte[] decoded = MimeTextCodec.DecodeTransfer(body, transferEncoding, diagnostics, location);
        if (decoded.LongLength > maximumBytes) {
            throw new InvalidDataException("The S/MIME payload exceeds the configured CMS limit.");
        }
        return decoded;
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

    internal static bool TryCanonicalizeLineEndings(
        byte[] source,
        long maximumBytes,
        out byte[] canonical) {
        canonical = source;
        long length = source.LongLength;
        bool changed = false;
        for (int index = 0; index < source.Length; index++) {
            if (source[index] == '\r') {
                if (index + 1 < source.Length && source[index + 1] == '\n') {
                    index++;
                } else {
                    length++;
                    changed = true;
                }
            } else if (source[index] == '\n') {
                length++;
                changed = true;
            }
            if (length > maximumBytes || length > int.MaxValue) return false;
        }
        if (!changed) return false;

        canonical = new byte[(int)length];
        int output = 0;
        for (int index = 0; index < source.Length; index++) {
            byte value = source[index];
            if (value == '\r') {
                canonical[output++] = (byte)'\r';
                if (index + 1 < source.Length && source[index + 1] == '\n') {
                    canonical[output++] = (byte)'\n';
                    index++;
                } else {
                    canonical[output++] = (byte)'\n';
                }
            } else if (value == '\n') {
                canonical[output++] = (byte)'\r';
                canonical[output++] = (byte)'\n';
            } else {
                canonical[output++] = value;
            }
        }
        return true;
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
