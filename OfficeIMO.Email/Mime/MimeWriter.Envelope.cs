namespace OfficeIMO.Email;

internal static partial class MimeWriter {
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
            if (ManagedHeaders.Contains(header.Name) || projectedMetadataNames.Contains(header.Name) || EmailTransportIntegrity.ShouldOmit(header.Name, options)) continue;
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

}
