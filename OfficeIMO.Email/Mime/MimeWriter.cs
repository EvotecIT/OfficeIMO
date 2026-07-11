namespace OfficeIMO.Email;

internal static class MimeWriter {
    private static readonly HashSet<string> ManagedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "Subject", "From", "Sender", "To", "Cc", "Bcc", "Reply-To", "Date", "Message-ID",
        "MIME-Version", "Content-Type", "Content-Transfer-Encoding", "Content-Disposition"
    };

    internal static byte[] Write(EmailDocument document, EmailWriterOptions options, IList<EmailDiagnostic> diagnostics) {
        MimeWriterState state = new MimeWriterState(options, diagnostics);
        using (MemoryStream output = new MemoryStream()) {
            WriteMessage(output, document, state, 0);
            return output.ToArray();
        }
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
        if (document.From != null) WriteLine(output, string.Concat("From: ", FormatAddress(document.From)));
        if (document.Sender != null) WriteLine(output, string.Concat("Sender: ", FormatAddress(document.Sender)));
        WriteRecipientHeader(output, document, EmailRecipientKind.To, "To");
        WriteRecipientHeader(output, document, EmailRecipientKind.Cc, "Cc");
        if (options.IncludeBccHeader) WriteRecipientHeader(output, document, EmailRecipientKind.Bcc, "Bcc");
        WriteRecipientHeader(output, document, EmailRecipientKind.ReplyTo, "Reply-To");
        if (document.Date.HasValue) {
            WriteLine(output, string.Concat("Date: ", document.Date.Value.ToString("ddd, dd MMM yyyy HH:mm:ss zzz", CultureInfo.InvariantCulture)));
        }
        if (!string.IsNullOrWhiteSpace(document.MessageId)) {
            WriteLine(output, string.Concat("Message-ID: <", SanitizeMessageId(document.MessageId!), ">"));
        }
        foreach (EmailHeader header in document.Headers) {
            if (ManagedHeaders.Contains(header.Name)) continue;
            WriteLine(output, string.Concat(SanitizeHeaderName(header.Name), ": ", EncodeHeaderText(header.Value)));
        }
    }

    private static void WriteRecipientHeader(Stream output, EmailDocument document, EmailRecipientKind kind, string name) {
        string[] addresses = document.Recipients.Where(item => item.Kind == kind).Select(item => FormatAddress(item.Address)).ToArray();
        if (addresses.Length > 0) WriteLine(output, string.Concat(name, ": ", string.Join(", ", addresses)));
    }

    private static void WriteContent(Stream output, EmailDocument document, MimeWriterState state, int depth, bool includeLeadingHeaders) {
        bool hasAttachments = document.Attachments.Count > 0;
        bool hasAlternative = document.Body.Text != null && document.Body.Html != null;
        if (hasAttachments) {
            string boundary = CreateBoundary(document, depth, "mixed");
            WriteLine(output, string.Concat("Content-Type: multipart/mixed; boundary=\"", boundary, "\""));
            WriteLine(output, string.Empty);
            WriteLine(output, string.Concat("--", boundary));
            WriteBodyEntity(output, document, state, depth, hasAlternative);
            for (int i = 0; i < document.Attachments.Count; i++) {
                WriteLine(output, string.Concat("--", boundary));
                WriteAttachment(output, document.Attachments[i], state, depth + 1, i);
            }
            WriteLine(output, string.Concat("--", boundary, "--"));
            return;
        }

        WriteBodyEntity(output, document, state, depth, hasAlternative);
    }

    private static void WriteBodyEntity(Stream output, EmailDocument document, MimeWriterState state, int depth, bool hasAlternative) {
        if (hasAlternative) {
            string boundary = CreateBoundary(document, depth, "alternative");
            WriteLine(output, string.Concat("Content-Type: multipart/alternative; boundary=\"", boundary, "\""));
            WriteLine(output, string.Empty);
            WriteLine(output, string.Concat("--", boundary));
            WriteTextPart(output, "text/plain", document.Body.Text!, state.Options.Base64LineLength);
            WriteLine(output, string.Concat("--", boundary));
            WriteTextPart(output, "text/html", document.Body.Html!, state.Options.Base64LineLength);
            WriteLine(output, string.Concat("--", boundary, "--"));
        } else if (document.Body.Html != null) {
            WriteTextPart(output, "text/html", document.Body.Html, state.Options.Base64LineLength);
        } else {
            WriteTextPart(output, "text/plain", document.Body.Text ?? string.Empty, state.Options.Base64LineLength);
        }
    }

    private static void WriteTextPart(Stream output, string mediaType, string text, int base64LineLength) {
        WriteLine(output, string.Concat("Content-Type: ", mediaType, "; charset=utf-8"));
        WriteLine(output, "Content-Transfer-Encoding: base64");
        WriteLine(output, string.Empty);
        WriteBase64(output, Encoding.UTF8.GetBytes(text), base64LineLength);
    }

    private static void WriteAttachment(Stream output, EmailAttachment attachment, MimeWriterState state, int depth, int index) {
        string contentType = string.IsNullOrWhiteSpace(attachment.ContentType)
            ? (attachment.EmbeddedDocument == null ? "application/octet-stream" : "message/rfc822")
            : attachment.ContentType!;
        string? fileName = attachment.FileName;
        WriteLine(output, string.Concat("Content-Type: ", SanitizeToken(contentType), FormatFileNameParameter("name", fileName)));
        WriteLine(output, string.Concat("Content-Disposition: ", attachment.IsInline ? "inline" : "attachment",
            FormatFileNameParameter("filename", fileName)));
        if (!string.IsNullOrWhiteSpace(attachment.ContentId)) {
            WriteLine(output, string.Concat("Content-ID: <", attachment.ContentId!.Trim().Trim('<', '>'), ">"));
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

        WriteLine(output, "Content-Transfer-Encoding: base64");
        WriteLine(output, string.Empty);
        byte[] content = attachment.Content ?? Array.Empty<byte>();
        if (attachment.Content == null && attachment.Length > 0) {
            state.Diagnostics.Add(new EmailDiagnostic("EMAIL_ATTACHMENT_CONTENT_UNAVAILABLE",
                string.Concat("Attachment ", index.ToString(CultureInfo.InvariantCulture),
                    " has a declared length but no retained content; an empty payload was written."),
                EmailDiagnosticSeverity.Error, string.Concat("attachment[", index.ToString(CultureInfo.InvariantCulture), "]")));
        }
        WriteBase64(output, content, state.Options.Base64LineLength);
    }

    private static void WriteBase64(Stream output, byte[] data, int lineLength) {
        string value = Convert.ToBase64String(data);
        if (value.Length == 0) {
            WriteLine(output, string.Empty);
            return;
        }
        for (int offset = 0; offset < value.Length; offset += lineLength) {
            int count = Math.Min(lineLength, value.Length - offset);
            WriteLine(output, value.Substring(offset, count));
        }
    }

    private static string CreateBoundary(EmailDocument document, int depth, string kind) {
        ulong hash = 14695981039346656037UL;
        Hash(ref hash, document.Subject);
        Hash(ref hash, document.MessageId);
        Hash(ref hash, document.Body.Text);
        Hash(ref hash, document.Body.Html);
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
        if (sanitized.Any(character => character < 32 || character > 126)) return EncodeHeaderText(sanitized);
        if (sanitized.All(IsAddressPhraseCharacter)) return sanitized;
        return string.Concat("\"", sanitized.Replace("\\", "\\\\").Replace("\"", "\\\""), "\"");
    }

    private static bool IsAddressPhraseCharacter(char character) {
        return (character >= 'A' && character <= 'Z') || (character >= 'a' && character <= 'z') ||
            (character >= '0' && character <= '9') || character == ' ' || character == '\t' ||
            "!#$%&'*+-/=?^_`{|}~".IndexOf(character) >= 0;
    }

    private static string EncodeHeaderText(string value) {
        string sanitized = value.Replace("\r", string.Empty).Replace("\n", " ");
        bool ascii = sanitized.All(character => character >= 32 && character <= 126);
        if (ascii) return sanitized;
        return string.Concat("=?utf-8?B?", Convert.ToBase64String(Encoding.UTF8.GetBytes(sanitized)), "?=");
    }

    private static string FormatFileNameParameter(string name, string? value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        string sanitized = value!.Replace("\r", string.Empty).Replace("\n", string.Empty);
        if (sanitized.All(character => character >= 32 && character <= 126) && sanitized.IndexOf('"') < 0) {
            return string.Concat("; ", name, "=\"", sanitized.Replace("\\", "\\\\"), "\"");
        }
        byte[] bytes = Encoding.UTF8.GetBytes(sanitized);
        StringBuilder encoded = new StringBuilder();
        for (int i = 0; i < bytes.Length; i++) {
            byte current = bytes[i];
            if ((current >= 'a' && current <= 'z') || (current >= 'A' && current <= 'Z') ||
                (current >= '0' && current <= '9') || current == '-' || current == '_' || current == '.') {
                encoded.Append((char)current);
            } else {
                encoded.Append('%').Append(current.ToString("X2", CultureInfo.InvariantCulture));
            }
        }
        return string.Concat("; ", name, "*=utf-8''", encoded.ToString());
    }

    private static string SanitizeHeaderName(string value) {
        StringBuilder result = new StringBuilder(value.Length);
        foreach (char character in value) {
            if ((character >= 'A' && character <= 'Z') || (character >= 'a' && character <= 'z') ||
                (character >= '0' && character <= '9') || character == '-') result.Append(character);
        }
        return result.Length == 0 ? "X-OfficeIMO-Header" : result.ToString();
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
        byte[] bytes = Encoding.ASCII.GetBytes(string.Concat(value, "\r\n"));
        output.Write(bytes, 0, bytes.Length);
    }
}
