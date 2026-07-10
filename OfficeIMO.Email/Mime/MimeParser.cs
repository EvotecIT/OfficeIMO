namespace OfficeIMO.Email;

internal static class MimeParser {
    internal static EmailDocument Parse(byte[] data, EmailReaderOptions options, IList<EmailDiagnostic> diagnostics) {
        MimeParserState state = new MimeParserState(options, diagnostics);
        return ParseMessage(data, 0, data.Length, state, 0, "message");
    }

    private static EmailDocument ParseMessage(byte[] data, int offset, int count, MimeParserState state,
        int nestedMessageDepth, string location) {
        EmailDocument document = new EmailDocument { Format = EmailFileFormat.Eml, OutlookItemKind = OutlookItemKind.Message };
        List<EmailHeader> headers = new List<EmailHeader>();
        int bodyOffset = MimeHeaderParser.Parse(data, offset, count, state.Options, headers, state.Diagnostics, location);
        foreach (EmailHeader header in headers) document.Headers.Add(header);
        PopulateEnvelope(document, headers, state.Diagnostics, location);

        int end = offset + count;
        int bodyCount = Math.Max(0, end - bodyOffset);
        ParseEntity(headers, data, bodyOffset, bodyCount, document, state, 0, nestedMessageDepth, location);
        return document;
    }

    private static void ParseEntity(IReadOnlyList<EmailHeader> headers, byte[] data, int offset, int count,
        EmailDocument document, MimeParserState state, int mimeDepth, int nestedMessageDepth, string location) {
        if (mimeDepth > state.Options.MaxMimeDepth) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxMimeDepth), mimeDepth, state.Options.MaxMimeDepth);
        }
        state.CountPart();

        MimeValue contentType = MimeValueParser.Parse(MimeHeaderParser.GetValue(headers, "Content-Type"),
            "text/plain", state.Diagnostics, location);
        MimeValue disposition = MimeValueParser.Parse(MimeHeaderParser.GetValue(headers, "Content-Disposition"),
            string.Empty, state.Diagnostics, location);
        string? transferEncoding = MimeHeaderParser.GetValue(headers, "Content-Transfer-Encoding");

        if (contentType.Value.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase)) {
            string? boundary = contentType.GetParameter("boundary");
            if (string.IsNullOrEmpty(boundary)) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BOUNDARY_MISSING",
                    string.Concat("Multipart entity '", contentType.Value, "' has no boundary."),
                    EmailDiagnosticSeverity.Error, location));
                return;
            }

            List<ArraySegment<byte>> parts = SplitMultipart(data, offset, count, boundary!, state.Diagnostics, location);
            for (int i = 0; i < parts.Count; i++) {
                string partLocation = string.Concat(location, "/part[", i.ToString(CultureInfo.InvariantCulture), "]");
                List<EmailHeader> partHeaders = new List<EmailHeader>();
                ArraySegment<byte> part = parts[i];
                int partBodyOffset = MimeHeaderParser.Parse(data, part.Offset, part.Count, state.Options,
                    partHeaders, state.Diagnostics, partLocation);
                int partEnd = part.Offset + part.Count;
                ParseEntity(partHeaders, data, partBodyOffset, Math.Max(0, partEnd - partBodyOffset),
                    document, state, mimeDepth + 1, nestedMessageDepth, partLocation);
            }
            return;
        }

        byte[] source = Copy(data, offset, count);
        byte[] decoded = MimeTextCodec.DecodeTransfer(source, transferEncoding, state.Diagnostics, location);
        string? fileName = disposition.GetParameter("filename") ?? contentType.GetParameter("name");
        bool attachmentDisposition = string.Equals(disposition.Value, "attachment", StringComparison.OrdinalIgnoreCase);
        bool inlineDisposition = string.Equals(disposition.Value, "inline", StringComparison.OrdinalIgnoreCase);
        bool isBody = !attachmentDisposition && string.IsNullOrWhiteSpace(fileName) &&
            (string.Equals(contentType.Value, "text/plain", StringComparison.OrdinalIgnoreCase) ||
             string.Equals(contentType.Value, "text/html", StringComparison.OrdinalIgnoreCase));

        if (isBody) {
            string? charset = contentType.GetParameter("charset");
            string text = MimeTextCodec.DecodeText(decoded, charset, state.Diagnostics, location);
            if (string.Equals(contentType.Value, "text/html", StringComparison.OrdinalIgnoreCase)) {
                if (document.Body.Html == null) {
                    document.Body.Html = text;
                    document.Body.HtmlCharset = charset;
                }
            } else if (document.Body.Text == null) {
                document.Body.Text = text;
                document.Body.TextCharset = charset;
            }
            return;
        }

        if (string.Equals(contentType.Value, "message/rfc822", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(contentType.Value, "message/global", StringComparison.OrdinalIgnoreCase)) {
            state.CountAttachment(decoded.LongLength);
            EmailAttachment embedded = CreateAttachment(headers, contentType, disposition, fileName, inlineDisposition, decoded, state.Options);
            if (nestedMessageDepth >= state.Options.MaxNestedMessageDepth) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_NESTED_MESSAGE_LIMIT",
                    "The embedded message was retained but not parsed because the nested-message limit was reached.",
                    EmailDiagnosticSeverity.Warning, location));
            } else {
                embedded.EmbeddedDocument = ParseMessage(decoded, 0, decoded.Length, state,
                    nestedMessageDepth + 1, string.Concat(location, "/message"));
            }
            document.Attachments.Add(embedded);
            return;
        }

        state.CountAttachment(decoded.LongLength);
        document.Attachments.Add(CreateAttachment(headers, contentType, disposition, fileName, inlineDisposition, decoded, state.Options));
    }

    private static EmailAttachment CreateAttachment(IReadOnlyList<EmailHeader> headers, MimeValue contentType,
        MimeValue disposition, string? fileName, bool inlineDisposition, byte[] decoded, EmailReaderOptions options) {
        return new EmailAttachment {
            FileName = fileName,
            ContentType = contentType.Value,
            ContentId = TrimAngleBrackets(MimeHeaderParser.GetValue(headers, "Content-ID")),
            ContentLocation = MimeHeaderParser.GetValue(headers, "Content-Location"),
            IsInline = inlineDisposition || !string.IsNullOrWhiteSpace(MimeHeaderParser.GetValue(headers, "Content-ID")),
            Length = decoded.LongLength,
            Content = options.IncludeAttachmentContent ? decoded : null
        };
    }

    private static void PopulateEnvelope(EmailDocument document, IReadOnlyList<EmailHeader> headers,
        IList<EmailDiagnostic> diagnostics, string location) {
        document.Subject = MimeHeaderParser.GetValue(headers, "Subject");
        document.From = MimeAddressParser.ParseOne(MimeHeaderParser.GetValue(headers, "From"));
        document.Sender = MimeAddressParser.ParseOne(MimeHeaderParser.GetValue(headers, "Sender"));
        document.MessageId = TrimAngleBrackets(MimeHeaderParser.GetValue(headers, "Message-ID"));
        document.Date = ParseDate(MimeHeaderParser.GetValue(headers, "Date"), diagnostics, string.Concat(location, "/Date"));

        AddRecipients(document, headers, "To", EmailRecipientKind.To);
        AddRecipients(document, headers, "Cc", EmailRecipientKind.Cc);
        AddRecipients(document, headers, "Bcc", EmailRecipientKind.Bcc);
        AddRecipients(document, headers, "Reply-To", EmailRecipientKind.ReplyTo);

        string? received = MimeHeaderParser.GetValues(headers, "Received").LastOrDefault();
        if (!string.IsNullOrWhiteSpace(received)) {
            int separator = received!.LastIndexOf(';');
            if (separator >= 0) {
                document.ReceivedDate = ParseDate(received.Substring(separator + 1), diagnostics,
                    string.Concat(location, "/Received"), false);
            }
        }
    }

    private static void AddRecipients(EmailDocument document, IEnumerable<EmailHeader> headers,
        string headerName, EmailRecipientKind kind) {
        foreach (string value in MimeHeaderParser.GetValues(headers, headerName)) {
            foreach (EmailAddress address in MimeAddressParser.ParseMany(value)) {
                document.Recipients.Add(new EmailRecipient(kind, address));
            }
        }
    }

    private static DateTimeOffset? ParseDate(string? value, IList<EmailDiagnostic> diagnostics, string location,
        bool reportInvalid = true) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        if (DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out DateTimeOffset result)) {
            return result;
        }
        if (reportInvalid) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_DATE_INVALID",
                string.Concat("Date value '", value, "' could not be parsed."), EmailDiagnosticSeverity.Warning, location));
        }
        return null;
    }

    private static List<ArraySegment<byte>> SplitMultipart(byte[] data, int offset, int count, string boundary,
        IList<EmailDiagnostic> diagnostics, string location) {
        byte[] marker = Encoding.ASCII.GetBytes(string.Concat("--", boundary));
        int end = offset + count;
        int lineStart = offset;
        int partStart = -1;
        bool closed = false;
        List<ArraySegment<byte>> parts = new List<ArraySegment<byte>>();

        while (lineStart < end) {
            int lineEnd = lineStart;
            while (lineEnd < end && data[lineEnd] != '\r' && data[lineEnd] != '\n') lineEnd++;
            if (IsBoundaryLine(data, lineStart, lineEnd, marker, out bool closing)) {
                if (partStart >= 0) {
                    int partEnd = TrimLineEndingBefore(data, partStart, lineStart);
                    parts.Add(new ArraySegment<byte>(data, partStart, Math.Max(0, partEnd - partStart)));
                }
                if (closing) {
                    closed = true;
                    break;
                }
                partStart = SkipLineEnding(data, lineEnd, end);
            }
            lineStart = SkipLineEnding(data, lineEnd, end);
        }

        if (!closed) {
            if (partStart >= 0 && partStart < end) {
                int partEnd = TrimLineEndingBefore(data, partStart, end);
                parts.Add(new ArraySegment<byte>(data, partStart, Math.Max(0, partEnd - partStart)));
            }
            diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BOUNDARY_NOT_CLOSED",
                string.Concat("Multipart boundary '", boundary, "' has no closing delimiter."),
                EmailDiagnosticSeverity.Warning, location));
        }
        if (partStart < 0) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BOUNDARY_NOT_FOUND",
                string.Concat("Multipart boundary '", boundary, "' was not found."),
                EmailDiagnosticSeverity.Error, location));
        }
        return parts;
    }

    private static bool IsBoundaryLine(byte[] data, int start, int end, byte[] marker, out bool closing) {
        closing = false;
        if (end - start < marker.Length) return false;
        for (int i = 0; i < marker.Length; i++) {
            if (data[start + i] != marker[i]) return false;
        }
        int position = start + marker.Length;
        if (position + 1 < end && data[position] == '-' && data[position + 1] == '-') {
            closing = true;
            position += 2;
        }
        while (position < end && (data[position] == ' ' || data[position] == '\t')) position++;
        return position == end;
    }

    private static int SkipLineEnding(byte[] data, int position, int end) {
        if (position < end && data[position] == '\r') position++;
        if (position < end && data[position] == '\n') position++;
        return position;
    }

    private static int TrimLineEndingBefore(byte[] data, int minimum, int position) {
        int result = position;
        if (result > minimum && data[result - 1] == '\n') result--;
        if (result > minimum && data[result - 1] == '\r') result--;
        return result;
    }

    private static byte[] Copy(byte[] data, int offset, int count) {
        if (offset == 0 && count == data.Length) return data;
        byte[] result = new byte[count];
        Buffer.BlockCopy(data, offset, result, 0, count);
        return result;
    }

    private static string? TrimAngleBrackets(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return value;
        string trimmed = value!.Trim();
        if (trimmed.Length >= 2 && trimmed[0] == '<' && trimmed[trimmed.Length - 1] == '>') {
            return trimmed.Substring(1, trimmed.Length - 2);
        }
        return trimmed;
    }
}
