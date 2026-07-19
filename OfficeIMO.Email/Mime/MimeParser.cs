namespace OfficeIMO.Email;

internal static class MimeParser {
    internal static EmailDocument Parse(byte[] data, EmailReaderOptions options, IList<EmailDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        MimeParserState state = new MimeParserState(options, diagnostics, cancellationToken);
        return ParseMessage(data, 0, data.Length, state, 0, "message");
    }

    private static EmailDocument ParseMessage(byte[] data, int offset, int count, MimeParserState state,
        int nestedMessageDepth, string location) {
        state.ThrowIfCancellationRequested();
        EmailDocument document = new EmailDocument { Format = EmailFileFormat.Eml, OutlookItemKind = OutlookItemKind.Message };
        List<EmailHeader> headers = new List<EmailHeader>();
        int bodyOffset = MimeHeaderParser.Parse(data, offset, count, state.Options, headers, state.Diagnostics, location);
        foreach (EmailHeader header in headers) document.Headers.Add(header);
        PopulateEnvelope(document, headers, state.Diagnostics, location);
        MimeMessageMetadataProjection.Apply(document, headers);

        int end = offset + count;
        int bodyCount = Math.Max(0, end - bodyOffset);
        ParseEntity(headers, data, bodyOffset, bodyCount, document, state, 0, nestedMessageDepth, location);
        MimeProtectionProjection.Apply(document, headers, state.Diagnostics, location);
        bool hasSemanticContent = document.Attachments.Any(attachment => attachment.IsProjectedSemanticContent);
        if (hasSemanticContent && document.MimeHasMessageBody && !document.MimeSemanticSourceHasTextBody &&
            !string.IsNullOrWhiteSpace(document.Body.Text)) {
            document.MimeSemanticProjectionIsIncomplete = true;
        }
        if (hasSemanticContent) {
            document.MimeSemanticSourceModelFingerprint = EmailDocumentStateFingerprint.TryCompute(document);
        }
        return document;
    }

    private static void ParseEntity(IReadOnlyList<EmailHeader> headers, byte[] data, int offset, int count,
        EmailDocument document, MimeParserState state, int mimeDepth, int nestedMessageDepth, string location,
        string defaultContentType = "text/plain", string? preferredBodyContentId = null,
        bool isRelatedSibling = false, bool isDefaultRelatedRoot = false) {
        if (mimeDepth > state.Options.MaxMimeDepth) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxMimeDepth), mimeDepth, state.Options.MaxMimeDepth);
        }
        state.CountPart();

        MimeValue contentType = MimeValueParser.Parse(MimeHeaderParser.GetValue(headers, "Content-Type"),
            defaultContentType, state.Diagnostics, location);
        MimeValue disposition = MimeValueParser.Parse(MimeHeaderParser.GetValue(headers, "Content-Disposition"),
            string.Empty, state.Diagnostics, location);
        string? transferEncoding = MimeHeaderParser.GetValue(headers, "Content-Transfer-Encoding");
        string? fileName = disposition.GetParameter("filename") ?? contentType.GetParameter("name");
        string? contentId = MimeHeaderParser.GetValue(headers, "Content-ID");
        string? contentLocation = MimeHeaderParser.GetValue(headers, "Content-Location");
        bool contentIdMatchesPreferred = !string.IsNullOrWhiteSpace(preferredBodyContentId) &&
            string.Equals(TrimAngleBrackets(contentId), TrimAngleBrackets(preferredBodyContentId),
                StringComparison.OrdinalIgnoreCase);
        bool isPreferredRelatedBody = isDefaultRelatedRoot || contentIdMatchesPreferred;
        bool hasRelatedIdentity = !string.IsNullOrWhiteSpace(contentId) ||
            !string.IsNullOrWhiteSpace(contentLocation);
        bool attachmentDisposition = string.Equals(disposition.Value, "attachment", StringComparison.OrdinalIgnoreCase);
        bool inlineDisposition = string.Equals(disposition.Value, "inline", StringComparison.OrdinalIgnoreCase);

        if (contentType.Value.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase) &&
            !attachmentDisposition && string.IsNullOrWhiteSpace(fileName) &&
            (!isRelatedSibling || !hasRelatedIdentity || isPreferredRelatedBody)) {
            string? boundary = contentType.GetParameter("boundary");
            if (boundary == null) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BOUNDARY_MISSING",
                    string.Concat("Multipart entity '", contentType.Value, "' has no boundary."),
                    EmailDiagnosticSeverity.Warning, location));
                return;
            }
            if (boundary.Length == 0) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BOUNDARY_EMPTY",
                    string.Concat("Multipart entity '", contentType.Value,
                        "' declares an empty boundary; compatible recovery was attempted."),
                    EmailDiagnosticSeverity.Warning, location));
            }

            List<ArraySegment<byte>> parts = SplitMultipart(data, offset, count, boundary!, state, location);
            string childDefaultContentType = string.Equals(contentType.Value, "multipart/digest", StringComparison.OrdinalIgnoreCase)
                ? "message/rfc822"
                : "text/plain";
            bool isRelated = string.Equals(contentType.Value, "multipart/related", StringComparison.OrdinalIgnoreCase);
            string? childPreferredBodyContentId = isRelated
                ? TrimAngleBrackets(contentType.GetParameter("start"))
                : preferredBodyContentId;
            for (int i = 0; i < parts.Count; i++) {
                state.ThrowIfCancellationRequested();
                string partLocation = string.Concat(location, "/part[", i.ToString(CultureInfo.InvariantCulture), "]");
                List<EmailHeader> partHeaders = new List<EmailHeader>();
                ArraySegment<byte> part = parts[i];
                int partBodyOffset = MimeHeaderParser.Parse(data, part.Offset, part.Count, state.Options,
                    partHeaders, state.Diagnostics, partLocation);
                int partEnd = part.Offset + part.Count;
                string? partPreferredBodyContentId = childPreferredBodyContentId;
                bool partIsDefaultRelatedRoot = isRelated && i == 0 &&
                    string.IsNullOrWhiteSpace(childPreferredBodyContentId);
                if (isRelated && string.IsNullOrWhiteSpace(partPreferredBodyContentId) && i == 0) {
                    partPreferredBodyContentId = TrimAngleBrackets(MimeHeaderParser.GetValue(partHeaders, "Content-ID"));
                }
                ParseEntity(partHeaders, data, partBodyOffset, Math.Max(0, partEnd - partBodyOffset),
                    document, state, mimeDepth + 1, nestedMessageDepth, partLocation, childDefaultContentType,
                    partPreferredBodyContentId, isRelated, partIsDefaultRelatedRoot);
            }
            return;
        }

        bool isBodyCandidate = !attachmentDisposition && string.IsNullOrWhiteSpace(fileName) &&
            (!isRelatedSibling || !hasRelatedIdentity || isPreferredRelatedBody) &&
            (string.Equals(contentType.Value, "text/plain", StringComparison.OrdinalIgnoreCase) ||
             string.Equals(contentType.Value, "text/html", StringComparison.OrdinalIgnoreCase) ||
             string.Equals(contentType.Value, "text/rtf", StringComparison.OrdinalIgnoreCase));
        bool bodySlotAvailable = string.Equals(contentType.Value, "text/plain", StringComparison.OrdinalIgnoreCase)
            ? document.Body.Text == null
            : string.Equals(contentType.Value, "text/html", StringComparison.OrdinalIgnoreCase)
                ? document.Body.Html == null
                : document.Body.Rtf == null;
        bool isBody = isBodyCandidate && bodySlotAvailable;
        bool additionalInlineBody = isBodyCandidate && !bodySlotAvailable;
        bool embeddedMessage = string.Equals(contentType.Value, "message/rfc822", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(contentType.Value, "message/global", StringComparison.OrdinalIgnoreCase);
        bool calendarContent = string.Equals(contentType.Value, "text/calendar", StringComparison.OrdinalIgnoreCase);
        bool vcardContent = VCardCodec.IsVCardContentType(contentType.Value, contentType.GetParameter("profile"));
        bool semanticBodyPart = (calendarContent || vcardContent) && !attachmentDisposition &&
            string.IsNullOrWhiteSpace(fileName) && string.IsNullOrWhiteSpace(contentId) &&
            string.IsNullOrWhiteSpace(contentLocation);
        bool skipAttachmentDecoding = !isBody && !state.Options.IncludeAttachmentContent && !embeddedMessage &&
            !semanticBodyPart;
        long decodedLength = isBody ? 0 : MimeTextCodec.GetDecodedLength(data, offset, count, transferEncoding,
            skipAttachmentDecoding ? state.Diagnostics : null, skipAttachmentDecoding ? location : null);
        if (!isBody) {
            state.EnsureAttachmentWithinLimits(decodedLength);
            if (skipAttachmentDecoding) {
                state.CountAttachment(decodedLength);
                document.Attachments.Add(CreateAttachment(headers, contentType, disposition, fileName,
                    inlineDisposition || additionalInlineBody, attachmentDisposition, null, decodedLength,
                    isRelatedSibling));
                return;
            }
        }

        byte[] source = Copy(data, offset, count);
        byte[] decoded = MimeTextCodec.DecodeTransfer(source, transferEncoding, state.Diagnostics, location);

        if (isBody) {
            document.MimeHasMessageBody = true;
            string? charset = contentType.GetParameter("charset");
            string text = MimeTextCodec.DecodeText(decoded, charset, state.Diagnostics, location);
            if (string.Equals(contentType.Value, "text/plain", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(contentType.GetParameter("format"), "flowed", StringComparison.OrdinalIgnoreCase)) {
                bool deleteSpace = string.Equals(contentType.GetParameter("delsp"), "yes",
                    StringComparison.OrdinalIgnoreCase);
                text = MimeFlowedTextCodec.Decode(text, deleteSpace);
            }
            if (string.Equals(contentType.Value, "text/rtf", StringComparison.OrdinalIgnoreCase)) {
                if (document.Body.Rtf == null) document.Body.Rtf = text;
            } else if (string.Equals(contentType.Value, "text/html", StringComparison.OrdinalIgnoreCase)) {
                if (document.Body.Html == null) {
                    document.Body.Html = text;
                    document.Body.HtmlCharset = charset;
                    document.Body.HtmlContentId = TrimAngleBrackets(contentId);
                    document.Body.HtmlContentLocation = contentLocation;
                    document.Body.IsHtmlRelatedRoot = isRelatedSibling && isPreferredRelatedBody;
                }
            } else if (document.Body.Text == null) {
                document.Body.Text = text;
                document.Body.TextCharset = charset;
            }
            return;
        }

        if (embeddedMessage) {
            state.CountAttachment(decoded.LongLength);
            EmailAttachment embedded = CreateAttachment(headers, contentType, disposition, fileName,
                inlineDisposition, attachmentDisposition, state.Options.IncludeAttachmentContent ? decoded : null,
                decoded.LongLength, isRelatedSibling);
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
        EmailAttachment attachment = CreateAttachment(headers, contentType, disposition, fileName,
            inlineDisposition || additionalInlineBody, attachmentDisposition,
            state.Options.IncludeAttachmentContent || semanticBodyPart ? decoded : null, decoded.LongLength,
            isRelatedSibling);
        if (semanticBodyPart) attachment.IsMimeBodyPart = true;
        string? semanticCharset = contentType.GetParameter("charset");
        int semanticDiagnosticStart = state.Diagnostics.Count;
        bool shouldProjectSemantic = attachment.IsMimeBodyPart && (calendarContent || vcardContent);
        string? semanticText = shouldProjectSemantic
            ? MimeTextCodec.DecodeText(decoded, semanticCharset, state.Diagnostics, location)
            : null;
        bool semanticCharsetFallback = state.Diagnostics.Skip(semanticDiagnosticStart).Any(diagnostic =>
            diagnostic.Code == "EMAIL_MIME_CHARSET_UNSUPPORTED");
        if (calendarContent && attachment.IsMimeBodyPart && IcsCalendarCodec.TryProject(
                semanticText!, document, state.Diagnostics, location, contentType.GetParameter("method"))) {
            attachment.IsProjectedSemanticContent = true;
        } else if (vcardContent && attachment.IsMimeBodyPart && VCardCodec.TryProject(
                       semanticText!, document, state.Diagnostics, location)) {
            attachment.IsProjectedSemanticContent = true;
        }
        if (attachment.IsProjectedSemanticContent && semanticCharsetFallback) {
            document.MimeSemanticProjectionIsIncomplete = true;
        }
        if (attachment.IsProjectedSemanticContent && vcardContent &&
            HasUnpreservedVCardContentType(contentType)) {
            document.MimeSemanticProjectionIsIncomplete = true;
        }
        if (attachment.IsProjectedSemanticContent && mimeDepth > 0 &&
            HasUnpreservedSemanticPartHeaders(headers)) {
            document.MimeSemanticProjectionIsIncomplete = true;
        }
        if (attachment.IsProjectedSemanticContent && document.Attachments.Any(existing =>
            existing.IsProjectedSemanticContent)) document.MimeSemanticProjectionIsIncomplete = true;
        document.Attachments.Add(attachment);
    }

    private static bool HasUnpreservedSemanticPartHeaders(IEnumerable<EmailHeader> headers) =>
        headers.Any(header =>
            !header.Name.Equals("Content-Type", StringComparison.OrdinalIgnoreCase) &&
            !header.Name.Equals("Content-Transfer-Encoding", StringComparison.OrdinalIgnoreCase));

    private static bool HasUnpreservedVCardContentType(MimeValue contentType) =>
        !contentType.Value.Equals("text/vcard", StringComparison.OrdinalIgnoreCase) ||
        contentType.Parameters.Keys.Any(parameter =>
            !parameter.Equals("charset", StringComparison.OrdinalIgnoreCase));

    internal static EmailAttachment CreateAttachment(IReadOnlyList<EmailHeader> headers, MimeValue contentType,
        MimeValue disposition, string? fileName, bool inlineDisposition, bool attachmentDisposition,
        byte[]? content, long length, bool isMimeRelated = false) {
        var attachment = new EmailAttachment {
            FileName = fileName,
            ContentType = contentType.Value,
            ContentId = TrimAngleBrackets(MimeHeaderParser.GetValue(headers, "Content-ID")),
            ContentLocation = MimeHeaderParser.GetValue(headers, "Content-Location"),
            IsInline = inlineDisposition ||
                !string.IsNullOrWhiteSpace(MimeHeaderParser.GetValue(headers, "Content-ID")) ||
                !string.IsNullOrWhiteSpace(MimeHeaderParser.GetValue(headers, "Content-Location")),
            IsMimeRelated = isMimeRelated,
            Length = length,
            Content = content,
            IsMimeAttachment = attachmentDisposition,
            IsMimeBodyPart = !attachmentDisposition && !inlineDisposition &&
                string.IsNullOrWhiteSpace(fileName) &&
                string.IsNullOrWhiteSpace(MimeHeaderParser.GetValue(headers, "Content-ID")) &&
                string.IsNullOrWhiteSpace(MimeHeaderParser.GetValue(headers, "Content-Location"))
        };
        foreach (KeyValuePair<string, string> parameter in contentType.Parameters) {
            if (!string.Equals(parameter.Key, "name", StringComparison.OrdinalIgnoreCase)) {
                attachment.ContentTypeParameters[parameter.Key] = parameter.Value;
            }
        }
        return attachment;
    }

    internal static void PopulateEnvelope(EmailDocument document, IReadOnlyList<EmailHeader> headers,
        IList<EmailDiagnostic> diagnostics, string location) {
        document.Subject = MimeHeaderParser.GetValue(headers, "Subject");
        document.From = MimeAddressParser.ParseOne(MimeHeaderParser.GetRawValue(headers, "From"), diagnostics,
            string.Concat(location, "/From"));
        document.Sender = MimeAddressParser.ParseOne(MimeHeaderParser.GetRawValue(headers, "Sender"), diagnostics,
            string.Concat(location, "/Sender"));
        document.MessageId = TrimAngleBrackets(MimeHeaderParser.GetValue(headers, "Message-ID"));
        document.Date = ParseDate(MimeHeaderParser.GetValue(headers, "Date"), diagnostics, string.Concat(location, "/Date"));

        AddRecipients(document, headers, "To", EmailRecipientKind.To, diagnostics, location);
        AddRecipients(document, headers, "Cc", EmailRecipientKind.Cc, diagnostics, location);
        AddRecipients(document, headers, "Bcc", EmailRecipientKind.Bcc, diagnostics, location);
        AddRecipients(document, headers, "Reply-To", EmailRecipientKind.ReplyTo, diagnostics, location);

        string? received = MimeHeaderParser.GetValues(headers, "Received").FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(received)) {
            int separator = received!.LastIndexOf(';');
            if (separator >= 0) {
                document.ReceivedDate = ParseDate(received.Substring(separator + 1), diagnostics,
                    string.Concat(location, "/Received"), false);
            }
        }
    }

    private static void AddRecipients(EmailDocument document, IEnumerable<EmailHeader> headers,
        string headerName, EmailRecipientKind kind, IList<EmailDiagnostic> diagnostics, string location) {
        foreach (string value in MimeHeaderParser.GetRawValues(headers, headerName)) {
            foreach (EmailAddress address in MimeAddressParser.ParseMany(value, diagnostics,
                string.Concat(location, "/", headerName))) {
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

    internal static List<ArraySegment<byte>> SplitMultipart(byte[] data, int offset, int count, string boundary,
        MimeParserState state, string location) {
        byte[] marker = Encoding.ASCII.GetBytes(string.Concat("--", boundary));
        int end = offset + count;
        int lineStart = offset;
        int partStart = -1;
        bool closed = false;
        List<ArraySegment<byte>> parts = new List<ArraySegment<byte>>();

        while (lineStart < end) {
            state.ThrowIfCancellationRequested();
            int lineEnd = lineStart;
            while (lineEnd < end && data[lineEnd] != '\r' && data[lineEnd] != '\n') lineEnd++;
            if (IsBoundaryLine(data, lineStart, lineEnd, marker, out bool closing)) {
                if (partStart >= 0) {
                    int partEnd = TrimLineEndingBefore(data, partStart, lineStart);
                    state.EnsurePendingPartCount(parts.Count + 1);
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
                state.EnsurePendingPartCount(parts.Count + 1);
                parts.Add(new ArraySegment<byte>(data, partStart, Math.Max(0, partEnd - partStart)));
            }
            state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BOUNDARY_NOT_CLOSED",
                string.Concat("Multipart boundary '", boundary, "' has no closing delimiter."),
                EmailDiagnosticSeverity.Warning, location));
        }
        if (partStart < 0) {
            state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_BOUNDARY_NOT_FOUND",
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

    internal static string? TrimAngleBrackets(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return value;
        string trimmed = value!.Trim();
        if (trimmed.Length >= 2 && trimmed[0] == '<' && trimmed[trimmed.Length - 1] == '>') {
            return trimmed.Substring(1, trimmed.Length - 2);
        }
        return trimmed;
    }
}
