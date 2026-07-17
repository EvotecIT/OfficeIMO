namespace OfficeIMO.Email;

internal static class MimeWriter {
    private static readonly HashSet<string> ManagedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "Subject", "From", "Sender", "To", "Cc", "Bcc", "Reply-To", "Date", "Message-ID",
        "References", "In-Reply-To", "MIME-Version", "Content-Type", "Content-Transfer-Encoding",
        "Content-Disposition"
    };

    internal static byte[] Write(EmailDocument document, EmailWriterOptions options, IList<EmailDiagnostic> diagnostics) {
        using (EmailBoundedMemoryStream output = new EmailBoundedMemoryStream(options.MaxOutputBytes)) {
            Write(output, document, options, diagnostics);
            return output.ToArray();
        }
    }

    internal static void Write(Stream output, EmailDocument document,
        EmailWriterOptions options, IList<EmailDiagnostic> diagnostics) {
        if (output == null || !output.CanWrite) {
            throw new ArgumentException("MIME output requires a writable stream.", nameof(output));
        }
        MimeWriterState state = new MimeWriterState(options, diagnostics);
        WriteMessage(output, document, state, 0);
    }

    private static void WriteMessage(Stream output, EmailDocument document, MimeWriterState state, int depth) {
        state.Enter(document, depth);
        try {
            WriteEnvelopeHeaders(output, document, state.Options);
            WriteLine(output, "MIME-Version: 1.0");
            WriteContent(output, document, state, depth, true);
        } finally {
            state.Exit(document);
        }
    }

    private static void WriteEnvelopeHeaders(Stream output, EmailDocument document, EmailWriterOptions options) {
        if (!string.IsNullOrWhiteSpace(document.Subject)) WriteLine(output, string.Concat("Subject: ", EncodeHeaderText(document.Subject!)));
        WriteAddressHeader(output, document, document.From, "From");
        WriteAddressHeader(output, document, document.Sender, "Sender");
        WriteRecipientHeader(output, document, EmailRecipientKind.To, "To");
        WriteRecipientHeader(output, document, EmailRecipientKind.Cc, "Cc");
        if (options.IncludeBccHeader) WriteRecipientHeader(output, document, EmailRecipientKind.Bcc, "Bcc");
        WriteRecipientHeader(output, document, EmailRecipientKind.ReplyTo, "Reply-To");
        if (document.Date.HasValue) {
            WriteLine(output, string.Concat("Date: ",
                document.Date.Value.ToString("ddd, dd MMM yyyy HH:mm:ss ", CultureInfo.InvariantCulture),
                FormatTimeZoneOffset(document.Date.Value.Offset)));
        }
        if (!string.IsNullOrWhiteSpace(document.MessageId)) {
            WriteLine(output, string.Concat("Message-ID: <", SanitizeMessageId(document.MessageId!), ">"));
        }
        WriteThreadingHeader(output, document, "References", document.MessageMetadata.InternetReferences);
        WriteThreadingHeader(output, document, "In-Reply-To", document.MessageMetadata.InReplyToId);
        EmailHeader[] projectedMetadataHeaders = MimeMessageMetadataProjection.CreateHeaders(document).ToArray();
        foreach (EmailHeader header in projectedMetadataHeaders) {
            WriteProjectedMetadataHeader(output, header);
        }
        var projectedMetadataNames = new HashSet<string>(
            projectedMetadataHeaders.Select(header => header.Name), StringComparer.OrdinalIgnoreCase);
        foreach (EmailHeader header in document.Headers) {
            if (ManagedHeaders.Contains(header.Name) || projectedMetadataNames.Contains(header.Name)) continue;
            WriteRetainedHeader(output, header);
        }
    }

    private static void WriteRetainedHeader(Stream output, EmailHeader header) {
        string name = MimeHeaderSafety.SanitizeName(header.Name);
        if (header.RawValue != null && header.RawValue.All(character =>
                character == '\t' || (character >= 32 && character <= 126))) {
            WriteFoldedRawHeader(output, name, header.RawValue);
            return;
        }

        WriteLine(output, string.Concat(name, ": ", EncodeHeaderText(header.Value)));
    }

    private static void WriteProjectedMetadataHeader(Stream output, EmailHeader header) {
        if (!string.Equals(header.Name, "Disposition-Notification-To", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(header.Name, "Return-Receipt-To", StringComparison.OrdinalIgnoreCase)) {
            WriteLine(output, string.Concat(header.Name, ": ", EncodeHeaderText(header.Value)));
            return;
        }

        var diagnostics = new List<EmailDiagnostic>();
        string[] addresses = MimeAddressParser.ParseMany(header.Value, diagnostics,
            string.Concat("transport/", header.Name)).Select(FormatAddress).ToArray();
        string value = addresses.Length == 0 ? EncodeHeaderText(header.Value) : string.Join(",\r\n ", addresses);
        WriteLine(output, string.Concat(header.Name, ": ", value));
    }

    private static void WriteFoldedRawHeader(Stream output, string name, string value) {
        string[] tokens = MimeHeaderSafety.SanitizeValue(value).Split(
            new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length == 0) {
            WriteLine(output, string.Concat(name, ":"));
            return;
        }

        var line = new StringBuilder(string.Concat(name, ":"));
        foreach (string token in tokens) {
            if (line.Length > name.Length + 1 && line.Length + token.Length + 1 > 78) {
                WriteLine(output, line.ToString());
                line.Clear().Append(' ');
            } else {
                line.Append(' ');
            }
            line.Append(token);
        }
        WriteLine(output, line.ToString());
    }

    private static void WriteRecipientHeader(Stream output, EmailDocument document, EmailRecipientKind kind, string name) {
        string[] addresses = document.Recipients.Where(item => item.Kind == kind)
            .Select(item => FormatAddress(item.Address)).ToArray();
        bool hasProjectedRecipients = document.Recipients.Any(item => item.Kind == kind);
        bool retainedHeaderIsParseable = false;
        if (addresses.Length == 0 && !hasProjectedRecipients) {
            var parsed = new List<EmailAddress>();
            var diagnostics = new List<EmailDiagnostic>();
            foreach (EmailHeader header in document.Headers.Where(header =>
                         string.Equals(header.Name, name, StringComparison.OrdinalIgnoreCase))) {
                parsed.AddRange(MimeAddressParser.ParseMany(header.RawValue ?? header.Value, diagnostics,
                    string.Concat("transport/", name)));
            }
            retainedHeaderIsParseable = parsed.Count > 0;
            addresses = parsed.Where(address => !HasProjectedRecipientAddress(document, address))
                .Select(FormatAddress).ToArray();
        }
        if (addresses.Length > 0) {
            WriteLine(output, string.Concat(name, ": ", string.Join(",\r\n ", addresses)));
        } else if (!hasProjectedRecipients && !retainedHeaderIsParseable) {
            foreach (EmailHeader header in document.Headers.Where(header =>
                         string.Equals(header.Name, name, StringComparison.OrdinalIgnoreCase))) {
                WriteRetainedHeader(output, header);
            }
        }
    }

    private static bool HasProjectedRecipientAddress(EmailDocument document, EmailAddress address) {
        string? candidate = address.Address?.Trim();
        return !string.IsNullOrWhiteSpace(candidate) && document.Recipients.Any(recipient =>
            string.Equals(recipient.Address.Address?.Trim(), candidate, StringComparison.OrdinalIgnoreCase));
    }

    private static void WriteAddressHeader(Stream output, EmailDocument document, EmailAddress? address, string name) {
        if (address != null) {
            WriteLine(output, string.Concat(name, ": ", FormatAddress(address)));
            return;
        }
        var diagnostics = new List<EmailDiagnostic>();
        EmailHeader? retained = document.Headers.FirstOrDefault(header =>
            string.Equals(header.Name, name, StringComparison.OrdinalIgnoreCase));
        EmailAddress? parsed = MimeAddressParser.ParseOne(retained?.RawValue ?? retained?.Value,
            diagnostics, string.Concat("transport/", name));
        if (parsed != null) WriteLine(output, string.Concat(name, ": ", FormatAddress(parsed)));
        else if (retained != null) WriteRetainedHeader(output, retained);
    }

    private static void WriteThreadingHeader(Stream output, EmailDocument document, string name, string? fallbackValue) {
        EmailHeader? retained = document.Headers.FirstOrDefault(header =>
            string.Equals(header.Name, name, StringComparison.OrdinalIgnoreCase));
        string? value = retained?.RawValue ?? retained?.Value ?? fallbackValue;
        if (string.IsNullOrWhiteSpace(value)) return;

        string[] tokens = MimeHeaderSafety.SanitizeValue(value!).Split(
            new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length == 0) return;
        var line = new StringBuilder(string.Concat(name, ":"));
        foreach (string token in tokens) {
            if (line.Length > name.Length + 1 && line.Length + token.Length + 1 > 78) {
                WriteLine(output, line.ToString());
                line.Clear().Append(' ');
            } else {
                line.Append(' ');
            }
            line.Append(token);
        }
        WriteLine(output, line.ToString());
    }

    private static void WriteContent(Stream output, EmailDocument document, MimeWriterState state, int depth, bool includeLeadingHeaders) {
        EmailAttachment? calendarAttachment = IcsCalendarCodec.FindSemanticAttachment(document);
        bool semanticSourceUnchanged = document.MimeSemanticSourceModelFingerprint != null &&
            EmailDocumentStateFingerprint.Matches(document, document.MimeSemanticSourceModelFingerprint);
        bool calendarIsAttachment = calendarAttachment != null &&
            IcsCalendarCodec.ShouldWriteAsAttachment(calendarAttachment);
        byte[]? retainedCalendarContent = calendarAttachment != null && semanticSourceUnchanged
            ? EmailAttachmentContent.ReadOrNull(calendarAttachment, state.Options.MaxOutputBytes)
            : null;
        bool calendarSourceReused = retainedCalendarContent != null;
        byte[]? calendarContent = document.OutlookItemKind == OutlookItemKind.Appointment ||
            document.OutlookItemKind == OutlookItemKind.Task
            ? calendarIsAttachment
                ? null
                : retainedCalendarContent ?? IcsCalendarCodec.Create(document)
            : null;
        EmailAttachment? vcardAttachment = VCardCodec.FindSemanticAttachment(document);
        var regularAttachmentList = document.Attachments.Where(attachment =>
            !ReferenceEquals(attachment, calendarAttachment) && !ReferenceEquals(attachment, vcardAttachment)).ToList();
        if (calendarIsAttachment && calendarAttachment != null) {
            regularAttachmentList.Add(calendarSourceReused
                ? calendarAttachment
                : IcsCalendarCodec.CreateRegeneratedAttachment(document, calendarAttachment));
        }
        EmailAttachment? contactBodyPart = null;
        if (document.OutlookItemKind == OutlookItemKind.Contact && document.Contact != null) {
            EmailAttachment contactPart = VCardCodec.CreateAttachment(document,
                semanticSourceUnchanged ? vcardAttachment : null);
            if (!document.MimeHasMessageBody && document.Body.Html == null && document.Body.Rtf == null &&
                calendarContent == null) {
                contactBodyPart = contactPart;
            } else {
                regularAttachmentList.Add(contactPart);
            }
        } else if (vcardAttachment != null) {
            regularAttachmentList.Add(vcardAttachment);
        }
        EmailAttachment[] regularAttachments = regularAttachmentList.ToArray();
        bool includeTextBody = document.Body.Text != null &&
            (document.MimeHasMessageBody || calendarAttachment == null && contactBodyPart == null &&
             document.OutlookItemKind != OutlookItemKind.Contact);
        bool hasAlternative = CountBodyAlternatives(document.Body, includeTextBody) +
            (calendarContent == null ? 0 : 1) > 1;
        bool hasRelatedResources = document.Body.IsHtmlRelatedRoot ||
            regularAttachments.Any(attachment => IsRelatedResource(document, attachment));
        bool hasUnrelatedAttachments = regularAttachments.Any(attachment => !IsRelatedResource(document, attachment));
        if (hasUnrelatedAttachments) {
            string boundary = CreateBoundary(document, depth, "mixed");
            WriteLine(output, string.Concat("Content-Type: multipart/mixed; boundary=\"", boundary, "\""));
            WriteLine(output, string.Empty);
            WriteLine(output, string.Concat("--", boundary));
            if (hasRelatedResources) {
                WriteRelatedBodyEntity(output, document, state, depth, hasAlternative, includeTextBody, calendarContent,
                    calendarSourceReused ? calendarAttachment : null, contactBodyPart, regularAttachments);
            } else {
                WriteBodyEntity(output, document, state, depth, hasAlternative, includeTextBody, calendarContent,
                    calendarSourceReused ? calendarAttachment : null, contactBodyPart);
            }
            for (int i = 0; i < regularAttachments.Length; i++) {
                if (IsRelatedResource(document, regularAttachments[i])) continue;
                WriteLine(output, string.Concat("--", boundary));
                WriteAttachment(output, regularAttachments[i], state, depth + 1, i);
            }
            WriteLine(output, string.Concat("--", boundary, "--"));
            return;
        }

        if (hasRelatedResources) {
            WriteRelatedBodyEntity(output, document, state, depth, hasAlternative, includeTextBody, calendarContent,
                calendarSourceReused ? calendarAttachment : null, contactBodyPart, regularAttachments);
        } else {
            WriteBodyEntity(output, document, state, depth, hasAlternative, includeTextBody, calendarContent,
                calendarSourceReused ? calendarAttachment : null, contactBodyPart);
        }
    }

    private static void WriteRelatedBodyEntity(Stream output, EmailDocument document, MimeWriterState state,
        int depth, bool hasAlternative, bool includeTextBody, byte[]? calendarContent,
        EmailAttachment? calendarAttachment,
        EmailAttachment? contactBodyPart,
        IReadOnlyList<EmailAttachment> attachments) {
        string boundary = CreateBoundary(document, depth, "related");
        string rootType = hasAlternative ? "multipart/alternative" : document.Body.Html != null
            ? "text/html"
            : "text/plain";
        string start = !hasAlternative && !string.IsNullOrWhiteSpace(document.Body.HtmlContentId)
            ? string.Concat("; start=\"<", SanitizeMessageId(document.Body.HtmlContentId!), ">\"")
            : string.Empty;
        WriteLine(output, string.Concat("Content-Type: multipart/related; boundary=\"", boundary,
            "\"; type=\"", rootType, "\"", start));
        WriteLine(output, string.Empty);
        WriteLine(output, string.Concat("--", boundary));
        WriteBodyEntity(output, document, state, depth, hasAlternative, includeTextBody, calendarContent,
            calendarAttachment, contactBodyPart);
        for (int i = 0; i < attachments.Count; i++) {
            if (!IsRelatedResource(document, attachments[i])) continue;
            WriteLine(output, string.Concat("--", boundary));
            WriteAttachment(output, attachments[i], state, depth + 1, i);
        }
        WriteLine(output, string.Concat("--", boundary, "--"));
    }

    private static bool IsRelatedResource(EmailDocument document, EmailAttachment attachment) {
        if (attachment.IsMimeRelated) return true;
        if (string.IsNullOrWhiteSpace(document.Body.Html)) return false;
        if (!string.IsNullOrWhiteSpace(attachment.ContentId) &&
            MimeRelatedResourceReference.ContainsContentId(document.Body.Html!, attachment.ContentId!)) {
            return true;
        }
        return !string.IsNullOrWhiteSpace(attachment.ContentLocation) &&
            MimeRelatedResourceReference.ContainsContentLocation(document.Body.Html!, attachment.ContentLocation!);
    }

    private static void WriteBodyEntity(Stream output, EmailDocument document, MimeWriterState state, int depth,
        bool hasAlternative, bool includeTextBody, byte[]? calendarContent, EmailAttachment? calendarAttachment,
        EmailAttachment? contactBodyPart) {
        if (hasAlternative) {
            string boundary = CreateBoundary(document, depth, "alternative");
            WriteLine(output, string.Concat("Content-Type: multipart/alternative; boundary=\"", boundary, "\""));
            WriteLine(output, string.Empty);
            if (includeTextBody) {
                WriteLine(output, string.Concat("--", boundary));
                WriteTextPart(output, "text/plain", document.Body.Text!, state.Options.Base64LineLength);
            }
            if (document.Body.Html != null) {
                WriteLine(output, string.Concat("--", boundary));
                WriteTextPart(output, "text/html", document.Body.Html, state.Options.Base64LineLength,
                    document.Body.HtmlContentId, document.Body.HtmlContentLocation);
            }
            if (document.Body.Rtf != null) {
                WriteLine(output, string.Concat("--", boundary));
                WriteRtfPart(output, document.Body.Rtf, state, "body/rtf");
            }
            if (calendarContent != null) {
                WriteLine(output, string.Concat("--", boundary));
                WriteCalendarPart(output, document, calendarContent, calendarAttachment, state.Options.Base64LineLength);
            }
            WriteLine(output, string.Concat("--", boundary, "--"));
        } else if (calendarContent != null) {
            WriteCalendarPart(output, document, calendarContent, calendarAttachment, state.Options.Base64LineLength);
        } else if (contactBodyPart != null) {
            WriteAttachment(output, contactBodyPart, state, depth + 1, 0);
        } else if (document.Body.Html != null) {
            WriteTextPart(output, "text/html", document.Body.Html, state.Options.Base64LineLength,
                document.Body.HtmlContentId, document.Body.HtmlContentLocation);
        } else if (document.Body.Rtf != null) {
            WriteRtfPart(output, document.Body.Rtf, state, "body/rtf");
        } else {
            WriteTextPart(output, "text/plain", document.Body.Text ?? string.Empty, state.Options.Base64LineLength);
        }
    }

    internal static string? CreateTransportHeaders(EmailDocument document, EmailWriterOptions options) {
        if (document.Headers.Count == 0) return null;
        using (var output = new MemoryStream()) {
            WriteTransportAddressHeader(output, document, "From", document.From, null);
            WriteTransportAddressHeader(output, document, "Sender", document.Sender, null);
            WriteTransportAddressHeader(output, document, "To", null, EmailRecipientKind.To);
            WriteTransportAddressHeader(output, document, "Cc", null, EmailRecipientKind.Cc);
            if (options.IncludeBccHeader || HasHeader(document, "Bcc")) {
                WriteTransportAddressHeader(output, document, "Bcc", null, EmailRecipientKind.Bcc);
            }
            WriteTransportAddressHeader(output, document, "Reply-To", null, EmailRecipientKind.ReplyTo);
            foreach (EmailHeader header in document.Headers) {
                if (IsAddressHeader(header.Name)) continue;
                WriteLine(output, string.Concat(MimeHeaderSafety.SanitizeName(header.Name), ": ",
                    MimeHeaderSafety.SanitizeValue(header.RawValue ?? header.Value)));
            }
            if (output.Length == 0) return null;
            return Encoding.UTF8.GetString(output.ToArray()).TrimEnd('\r', '\n');
        }
    }

    private static void WriteTransportAddressHeader(Stream output, EmailDocument document, string name,
        EmailAddress? scalarAddress, EmailRecipientKind? recipientKind) {
        bool retained = HasHeader(document, name);
        bool structured = scalarAddress != null || recipientKind.HasValue &&
            document.Recipients.Any(recipient => recipient.Kind == recipientKind.Value);
        if (!retained && !structured) return;
        if (recipientKind.HasValue) WriteRecipientHeader(output, document, recipientKind.Value, name);
        else WriteAddressHeader(output, document, scalarAddress, name);
    }

    private static bool HasHeader(EmailDocument document, string name) => document.Headers.Any(header =>
        string.Equals(header.Name, name, StringComparison.OrdinalIgnoreCase));

    private static bool IsAddressHeader(string name) =>
        string.Equals(name, "From", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "Sender", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "To", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "Cc", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "Bcc", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "Reply-To", StringComparison.OrdinalIgnoreCase);

    private static void WriteCalendarPart(Stream output, EmailDocument document, byte[] content,
        EmailAttachment? source, int base64LineLength) {
        string parameters = source == null
            ? string.Concat("; method=", SanitizeToken(IcsCalendarCodec.GetMethod(document)), "; charset=utf-8")
            : FormatContentTypeParameters(source.ContentTypeParameters);
        WriteLine(output, string.Concat("Content-Type: text/calendar", parameters));
        WriteLine(output, "Content-Transfer-Encoding: base64");
        WriteLine(output, string.Empty);
        WriteBase64(output, content, base64LineLength);
    }

    private static int CountBodyAlternatives(EmailBody body, bool includeTextBody) {
        return (includeTextBody ? 1 : 0) + (body.Html == null ? 0 : 1) + (body.Rtf == null ? 0 : 1);
    }

    private static void WriteTextPart(Stream output, string mediaType, string text, int base64LineLength,
        string? contentId = null, string? contentLocation = null) {
        WriteLine(output, string.Concat("Content-Type: ", mediaType, "; charset=utf-8"));
        WriteLine(output, "Content-Transfer-Encoding: base64");
        if (!string.IsNullOrWhiteSpace(contentId)) {
            WriteLine(output, string.Concat("Content-ID: <", SanitizeMessageId(contentId!), ">"));
        }
        if (!string.IsNullOrWhiteSpace(contentLocation)) {
            WriteLine(output, string.Concat("Content-Location: ", EncodeHeaderText(contentLocation!)));
        }
        WriteLine(output, string.Empty);
        WriteBase64(output, Encoding.UTF8.GetBytes(text), base64LineLength);
    }

    private static void WriteRtfPart(Stream output, string rtf, MimeWriterState state, string location) {
        WriteLine(output, "Content-Type: text/rtf; charset=iso-8859-1");
        WriteLine(output, "Content-Transfer-Encoding: base64");
        WriteLine(output, string.Empty);
        if (!EmailRtfByteCodec.TryEncode(rtf, out byte[] rtfBytes)) {
            state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MIME_RTF_CHARACTER_UNENCODABLE",
                "The RTF source contains a character above U+00FF. Serialize it through OfficeIMO.Rtf so the character is represented by an RTF escape.",
                EmailDiagnosticSeverity.Error, location));
        }
        WriteBase64(output, rtfBytes, state.Options.Base64LineLength);
    }

    private static void WriteAttachment(Stream output, EmailAttachment attachment, MimeWriterState state, int depth, int index) {
        bool embeddedMessage = attachment.EmbeddedDocument != null;
        string contentType = embeddedMessage
            ? "message/rfc822"
            : string.IsNullOrWhiteSpace(attachment.ContentType) ? "application/octet-stream" : attachment.ContentType!;
        if (!embeddedMessage && contentType.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase)) {
            if (!attachment.ContentTypeParameters.TryGetValue("boundary", out string? retainedBoundary) ||
                string.IsNullOrWhiteSpace(retainedBoundary)) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MULTIPART_ATTACHMENT_WRITTEN_OPAQUE",
                    "A retained multipart attachment was written as application/octet-stream because its boundary metadata is unavailable.",
                    EmailDiagnosticSeverity.Warning, string.Concat("attachment[", index.ToString(CultureInfo.InvariantCulture), "]")));
                contentType = "application/octet-stream";
            }
        }
        string? fileName = attachment.FileName;
        WriteLine(output, string.Concat("Content-Type: ", SanitizeToken(contentType),
            FormatContentTypeParameters(attachment.ContentTypeParameters), FormatFileNameParameter("name", fileName)));
        if (!attachment.IsMimeBodyPart) {
            WriteLine(output, string.Concat("Content-Disposition: ",
                attachment.IsMimeAttachment ? "attachment" : attachment.IsInline ? "inline" : "attachment",
                FormatFileNameParameter("filename", fileName)));
        }
        if (!string.IsNullOrWhiteSpace(attachment.ContentId)) {
            WriteLine(output, string.Concat("Content-ID: <", SanitizeMessageId(attachment.ContentId!), ">"));
        }
        if (!string.IsNullOrWhiteSpace(attachment.ContentLocation)) {
            WriteLine(output, string.Concat("Content-Location: ", EncodeHeaderText(attachment.ContentLocation!)));
        }

        if (attachment.EmbeddedDocument != null) {
            WriteLine(output, "Content-Transfer-Encoding: 8bit");
            WriteLine(output, string.Empty);
            WriteMessage(output, attachment.EmbeddedDocument, state, depth);
            return;
        }

        bool hasContent = attachment.Content != null || attachment.ContentSource != null ||
            EmailAttachmentStreamScope.HasStagedContent(attachment);
        if (!hasContent && attachment.Length > 0) {
            state.Diagnostics.Add(new EmailDiagnostic("EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE",
                string.Concat("Attachment ", index.ToString(CultureInfo.InvariantCulture),
                    " has a declared length but no retained content; an empty payload was written."),
                EmailDiagnosticSeverity.Error, string.Concat("attachment[", index.ToString(CultureInfo.InvariantCulture), "]")));
        }
        if (contentType.StartsWith("multipart/", StringComparison.OrdinalIgnoreCase)) {
            WriteLine(output, "Content-Transfer-Encoding: 8bit");
            WriteLine(output, string.Empty);
            using (Stream input = EmailAttachmentStreamScope.OpenRead(attachment)) {
                WriteRawEntity(output, input, state.Options.MaxOutputBytes);
            }
            return;
        }

        WriteLine(output, "Content-Transfer-Encoding: base64");
        WriteLine(output, string.Empty);
        using (Stream input = EmailAttachmentStreamScope.OpenRead(attachment)) {
            WriteBase64(output, input, state.Options.Base64LineLength,
                state.Options.MaxOutputBytes);
        }
    }

    private static void WriteBase64(Stream output, byte[] data, int lineLength) {
        using (var input = new MemoryStream(data, writable: false)) {
            WriteBase64(output, input, lineLength, data.LongLength);
        }
    }

    private static void WriteBase64(Stream output, Stream input, int lineLength, long maximumInputBytes) {
        int bytesPerLine = checked(lineLength / 4 * 3);
        var buffer = new byte[bytesPerLine];
        long total = 0;
        bool wrote = false;
        while (true) {
            int count = 0;
            while (count < buffer.Length) {
                int read = input.Read(buffer, count, buffer.Length - count);
                if (read == 0) break;
                count += read;
                total = checked(total + read);
                if (total > maximumInputBytes) {
                    throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes),
                        total, maximumInputBytes);
                }
            }
            if (count == 0) break;
            WriteLine(output, Convert.ToBase64String(buffer, 0, count));
            wrote = true;
            if (count < buffer.Length) break;
        }
        if (!wrote) {
            WriteLine(output, string.Empty);
        }
    }

    private static string CreateBoundary(EmailDocument document, int depth, string kind) {
        ulong hash = 14695981039346656037UL;
        Hash(ref hash, document.Subject);
        Hash(ref hash, document.MessageId);
        Hash(ref hash, document.Body.Text);
        Hash(ref hash, document.Body.Html);
        Hash(ref hash, document.Body.Rtf);
        Hash(ref hash, document.Attachments.Count.ToString(CultureInfo.InvariantCulture));
        Hash(ref hash, depth.ToString(CultureInfo.InvariantCulture));
        Hash(ref hash, kind);
        return string.Concat("=_OfficeIMO_", kind, "_", hash.ToString("x16", CultureInfo.InvariantCulture));
    }

    private static void Hash(ref ulong hash, string? value) {
        if (value == null) return;
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        for (int i = 0; i < bytes.Length; i++) {
            hash ^= bytes[i];
            hash *= 1099511628211UL;
        }
    }

    private static string FormatAddress(EmailAddress address) {
        string value = SanitizeAddress(address.Address ?? address.RawValue ?? string.Empty);
        if (string.IsNullOrWhiteSpace(address.DisplayName)) return value;
        return string.Concat(FormatDisplayName(address.DisplayName!), " <", value.Trim('<', '>'), ">");
    }

    private static string FormatDisplayName(string value) {
        string sanitized = value.Replace("\r", string.Empty).Replace("\n", " ");
        if (sanitized.Length > 72 || sanitized.Any(character => character < 32 || character > 126)) return EncodeHeaderText(sanitized);
        if (sanitized.All(IsAddressPhraseCharacter)) return sanitized;
        return string.Concat("\"", sanitized.Replace("\\", "\\\\").Replace("\"", "\\\""), "\"");
    }

    private static bool IsAddressPhraseCharacter(char character) {
        return (character >= 'A' && character <= 'Z') || (character >= 'a' && character <= 'z') ||
            (character >= '0' && character <= '9') || character == ' ' || character == '\t' ||
            "!#$%&'*+-/=?^_`{|}~".IndexOf(character) >= 0;
    }

    private static string EncodeHeaderText(string value) {
        string sanitized = MimeHeaderSafety.SanitizeValue(value);
        bool ascii = sanitized.All(character => character >= 32 && character <= 126);
        if (ascii && sanitized.Length <= 72) return sanitized;
        const int maxEncodedBytes = 45;
        var words = new List<string>();
        var chunk = new StringBuilder();
        int chunkBytes = 0;
        for (int index = 0; index < sanitized.Length;) {
            int characterLength = char.IsHighSurrogate(sanitized[index]) && index + 1 < sanitized.Length &&
                char.IsLowSurrogate(sanitized[index + 1]) ? 2 : 1;
            int characterBytes = Encoding.UTF8.GetByteCount(sanitized.Substring(index, characterLength));
            if (chunkBytes > 0 && chunkBytes + characterBytes > maxEncodedBytes) {
                words.Add(EncodeWord(chunk.ToString()));
                chunk.Clear();
                chunkBytes = 0;
            }
            chunk.Append(sanitized, index, characterLength);
            chunkBytes += characterBytes;
            index += characterLength;
        }
        if (chunk.Length > 0) words.Add(EncodeWord(chunk.ToString()));
        return string.Join("\r\n ", words);
    }

    private static string EncodeWord(string value) {
        return string.Concat("=?utf-8?B?", Convert.ToBase64String(Encoding.UTF8.GetBytes(value)), "?=");
    }

    private static string FormatTimeZoneOffset(TimeSpan offset) {
        int totalMinutes = checked((int)offset.TotalMinutes);
        char sign = totalMinutes < 0 ? '-' : '+';
        int absoluteMinutes = Math.Abs(totalMinutes);
        return string.Concat(sign.ToString(), (absoluteMinutes / 60).ToString("00", CultureInfo.InvariantCulture),
            (absoluteMinutes % 60).ToString("00", CultureInfo.InvariantCulture));
    }

    private static string FormatFileNameParameter(string name, string? value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        string sanitized = value!.Replace("\r", string.Empty).Replace("\n", string.Empty);
        if (sanitized.Length <= 60 && sanitized.All(character => character >= 32 && character <= 126) &&
            sanitized.IndexOf('"') < 0) {
            return string.Concat("; ", name, "=\"", sanitized.Replace("\\", "\\\\"), "\"");
        }
        byte[] bytes = Encoding.UTF8.GetBytes(sanitized);
        var tokens = new List<string>(bytes.Length);
        for (int i = 0; i < bytes.Length; i++) {
            byte current = bytes[i];
            if ((current >= 'a' && current <= 'z') || (current >= 'A' && current <= 'Z') ||
                (current >= '0' && current <= '9') || current == '-' || current == '_' || current == '.') {
                tokens.Add(((char)current).ToString());
            } else {
                tokens.Add(string.Concat("%", current.ToString("X2", CultureInfo.InvariantCulture)));
            }
        }
        string encoded = string.Concat(tokens);
        if (encoded.Length <= 60) return string.Concat("; ", name, "*=utf-8''", encoded);

        var result = new StringBuilder();
        var segment = new StringBuilder();
        int segmentIndex = 0;
        foreach (string token in tokens) {
            if (segment.Length > 0 && segment.Length + token.Length > 45) {
                AppendParameterContinuation(result, name, segmentIndex++, segment.ToString());
                segment.Clear();
            }
            segment.Append(token);
        }
        if (segment.Length > 0) AppendParameterContinuation(result, name, segmentIndex, segment.ToString());
        return result.ToString();
    }

    private static void AppendParameterContinuation(StringBuilder output, string name, int index, string value) {
        output.Append(";\r\n ").Append(name).Append('*').Append(index.ToString(CultureInfo.InvariantCulture))
            .Append("*=");
        if (index == 0) output.Append("utf-8''");
        output.Append(value);
    }

    private static string FormatContentTypeParameters(IEnumerable<KeyValuePair<string, string>> parameters) {
        var result = new StringBuilder();
        foreach (KeyValuePair<string, string> parameter in parameters
            .Where(parameter => !string.Equals(parameter.Key, "name", StringComparison.OrdinalIgnoreCase))
            .OrderBy(parameter => parameter.Key, StringComparer.OrdinalIgnoreCase)) {
            string name = new string(parameter.Key.Where(character => char.IsLetterOrDigit(character) ||
                character == '-' || character == '_').ToArray());
            if (name.Length > 0) result.Append(FormatFileNameParameter(name, parameter.Value));
        }
        return result.ToString();
    }

    private static void WriteRawEntity(Stream output, Stream input, long maximumInputBytes) {
        var buffer = new byte[81920];
        long total = 0;
        int trailingFirst = -1;
        int trailingSecond = -1;
        while (true) {
            int read = input.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            total = checked(total + read);
            if (total > maximumInputBytes) {
                throw new EmailLimitExceededException(nameof(EmailWriterOptions.MaxOutputBytes),
                    total, maximumInputBytes);
            }
            output.Write(buffer, 0, read);
            if (read == 1) {
                trailingFirst = trailingSecond;
                trailingSecond = buffer[0];
            } else {
                trailingFirst = buffer[read - 2];
                trailingSecond = buffer[read - 1];
            }
        }
        if (trailingFirst != '\r' || trailingSecond != '\n') {
            WriteLine(output, string.Empty);
        }
    }

    private static string SanitizeAddress(string value) {
        return value.Replace("\r", string.Empty).Replace("\n", string.Empty).Trim();
    }

    private static string SanitizeMessageId(string value) {
        return value.Replace("\r", string.Empty).Replace("\n", string.Empty).Trim().Trim('<', '>');
    }

    private static string SanitizeToken(string value) {
        return value.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace(";", string.Empty).Trim();
    }

    private static void WriteLine(Stream output, string value) {
        byte[] bytes = Encoding.UTF8.GetBytes(string.Concat(value, "\r\n"));
        output.Write(bytes, 0, bytes.Length);
    }
}
