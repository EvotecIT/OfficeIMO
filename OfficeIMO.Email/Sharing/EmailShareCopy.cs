namespace OfficeIMO.Email;

/// <summary>Creates an independent field-selected share artifact. Retained body and payload content still require caller review.</summary>
public static class EmailShareCopy {
    /// <summary>Copies only explicitly selected metadata and attachments, with a plain-text body and omission evidence.</summary>
    public static EmailShareCopyResult Create(EmailDocument source, EmailShareCopyOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        var effective = options ?? new EmailShareCopyOptions();
        Validate(effective);
        cancellationToken.ThrowIfCancellationRequested();
        if (source.Headers.Count > effective.MaxHeaders) throw new EmailLimitExceededException(nameof(effective.MaxHeaders), source.Headers.Count, effective.MaxHeaders);
        if (source.Recipients.Count > effective.MaxRecipients) throw new EmailLimitExceededException(nameof(effective.MaxRecipients), source.Recipients.Count, effective.MaxRecipients);
        string? body = effective.ReplacementBodyText ?? source.Body.Text;
        CheckText(body);
        string? subject = effective.ReplacementSubject ?? (Has(EmailShareFields.Subject) ? source.Subject : null);
        CheckText(subject);
        var diagnostics = new List<EmailDiagnostic>();
        var changes = new List<EmailShareFieldChange>();
        var copy = new EmailDocument { Format = EmailFileFormat.Eml, Subject = subject,
            Date = Has(EmailShareFields.Date) ? source.Date : null };
        copy.Body.Text = body;
        copy.Body.TextCharset = "utf-8";
        changes.Add(new EmailShareFieldChange("subject", effective.ReplacementSubject != null ? "replaced" : Has(EmailShareFields.Subject) ? "copied" : "omitted"));
        changes.Add(new EmailShareFieldChange("date", Has(EmailShareFields.Date) ? "copied" : "omitted"));
        changes.Add(new EmailShareFieldChange("body", effective.ReplacementBodyText != null ? "replaced" : body != null ? "copied" : "omitted"));
        changes.Add(new EmailShareFieldChange("body/html-rtf", "omitted"));
        changes.Add(new EmailShareFieldChange("source/raw-mapi-tnef-protection", "omitted"));
        changes.Add(new EmailShareFieldChange("threading", "omitted"));
        if (effective.ReplacementBodyText == null && body == null && (!string.IsNullOrEmpty(source.Body.Html) || !string.IsNullOrEmpty(source.Body.Rtf)))
            diagnostics.Add(new EmailDiagnostic("EMAIL_SHARE_PLAIN_BODY_UNAVAILABLE", "No plain body was copied. Supply replacement text or use the optional HTML indexing projection."));
        if (Has(EmailShareFields.Author) && source.From != null) copy.From = CopyAddress(source.From, "from");
        else changes.Add(new EmailShareFieldChange("from", "omitted"));
        for (int index = 0; index < source.Recipients.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            var recipient = source.Recipients[index];
            string path = "recipients/" + index.ToString(CultureInfo.InvariantCulture);
            if (!Has(EmailShareFields.Recipients) || recipient.Kind != EmailRecipientKind.To && recipient.Kind != EmailRecipientKind.Cc) {
                changes.Add(new EmailShareFieldChange(path, "omitted"));
                continue;
            }
            var address = CopyAddress(recipient.Address, path);
            if (address != null) copy.Recipients.Add(new EmailRecipient(recipient.Kind, address));
        }
        var names = new HashSet<string>(effective.RetainedHeaderNames, StringComparer.OrdinalIgnoreCase);
        bool integrity = source.Protection.IsProtected;
        for (int index = 0; index < source.Headers.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            var header = source.Headers[index];
            bool sensitiveIntegrity = EmailTransportIntegrity.IsSignatureChain(header.Name) || EmailTransportIntegrity.IsPayloadDependent(header.Name);
            integrity |= sensitiveIntegrity;
            if (!names.Contains(header.Name)) continue;
            bool retain = !sensitiveIntegrity && !IsOwnedHeader(header.Name);
            changes.Add(new EmailShareFieldChange("headers/" + index.ToString(CultureInfo.InvariantCulture), retain ? "copied" : "omitted"));
            if (retain) { CheckText(header.Value); copy.Headers.Add(new EmailHeader(header.Name, header.Value)); }
        }
        changes.Add(new EmailShareFieldChange("headers/unselected", "omitted"));
        if (integrity) diagnostics.Add(new EmailDiagnostic("EMAIL_SHARE_INTEGRITY_REMOVED", "Original signatures, payload digests and protected wrappers were not inherited. The new artifact has no original integrity assurance."));
        if (copy.Headers.Count > 0) diagnostics.Add(new EmailDiagnostic("EMAIL_SHARE_HEADER_RETAINED", "Explicitly selected extra headers are unchanged and may contain private information."));
        long bytes = 0;
        var selected = new HashSet<int>();
        foreach (int index in effective.AttachmentIndexes) {
            cancellationToken.ThrowIfCancellationRequested();
            if (index < 0 || index >= source.Attachments.Count) throw new ArgumentOutOfRangeException(nameof(options), "A selected attachment index is outside the source collection.");
            if (!selected.Add(index)) continue;
            var attachment = source.Attachments[index];
            if (attachment.EmbeddedDocument != null) throw new ArgumentException("Embedded messages require separate field-selected copies; they cannot be inherited as share attachments.", nameof(options));
            byte[]? content = EmailAttachmentContent.ReadOrNull(attachment,
                Math.Min(effective.MaxAttachmentBytes, effective.MaxTotalAttachmentBytes - bytes), cancellationToken);
            if (content == null) throw new InvalidDataException("Selected attachment content is unavailable. Linked paths are never opened.");
            bytes = checked(bytes + content.LongLength);
            string? name = effective.AttachmentNames.TryGetValue(index, out string? replacement) ? replacement
                : effective.KeepAttachmentNames ? attachment.FileName : "attachment-" + index.ToString("D4", CultureInfo.InvariantCulture) + ".bin";
            CheckText(name);
            CheckText(attachment.ContentType);
            copy.Attachments.Add(new EmailAttachment { FileName = name, ContentType = attachment.ContentType ?? "application/octet-stream",
                Content = ReferenceEquals(content, attachment.Content) ? (byte[])content.Clone() : content, Length = content.LongLength });
            changes.Add(new EmailShareFieldChange("attachments/" + index.ToString(CultureInfo.InvariantCulture), "copied"));
            changes.Add(new EmailShareFieldChange("attachments/" + index.ToString(CultureInfo.InvariantCulture) + "/name",
                effective.KeepAttachmentNames && !effective.AttachmentNames.ContainsKey(index) ? "copied" : "replaced"));
        }
        changes.Add(new EmailShareFieldChange("attachments/unselected", "omitted"));
        if (selected.Count > 0) diagnostics.Add(new EmailDiagnostic("EMAIL_SHARE_PAYLOAD_RETAINED", "Selected attachment payload bytes are unchanged and may contain private information; only their surrounding metadata was selected."));
        cancellationToken.ThrowIfCancellationRequested();
        return new EmailShareCopyResult(copy, changes.AsReadOnly(), diagnostics.AsReadOnly());

        bool Has(EmailShareFields field) => (effective.RetainedFields & field) != 0;
        void CheckText(string? value) {
            if (value?.Length > effective.MaxTextChars) throw new EmailLimitExceededException(nameof(effective.MaxTextChars), value.Length, effective.MaxTextChars);
        }
        EmailAddress? CopyAddress(EmailAddress address, string path) {
            string? value = address.Address;
            bool replaced = value != null && effective.AddressReplacements.TryGetValue(value, out _);
            if (replaced) value = effective.AddressReplacements[value!];
            CheckText(value);
            string? normalized = EmailSmtpAddress.Normalize(new EmailAddress(value) { AddressType = replaced ? "SMTP" : address.AddressType });
            if (normalized == null) {
                changes.Add(new EmailShareFieldChange(path, "omitted"));
                diagnostics.Add(new EmailDiagnostic("EMAIL_SHARE_ADDRESS_OMITTED", "The selected address was removed or could not be represented as SMTP.", location: path));
                return null;
            }
            if (effective.KeepDisplayNames) CheckText(address.DisplayName);
            changes.Add(new EmailShareFieldChange(path, replaced ? "replaced" : "copied"));
            return new EmailAddress(normalized, effective.KeepDisplayNames ? address.DisplayName : null) { AddressType = "SMTP" };
        }
    }

    private static bool IsOwnedHeader(string name) => name.StartsWith("Content-", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("Resent-", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("MIME-Version", StringComparison.OrdinalIgnoreCase) || name.Equals("From", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("Sender", StringComparison.OrdinalIgnoreCase) || name.Equals("To", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("Cc", StringComparison.OrdinalIgnoreCase) || name.Equals("Bcc", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("Reply-To", StringComparison.OrdinalIgnoreCase) || name.Equals("Subject", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("Date", StringComparison.OrdinalIgnoreCase) || name.Equals("Message-ID", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("References", StringComparison.OrdinalIgnoreCase) || name.Equals("In-Reply-To", StringComparison.OrdinalIgnoreCase);

    private static void Validate(EmailShareCopyOptions options) {
        const EmailShareFields known = EmailShareFields.Subject | EmailShareFields.Author | EmailShareFields.Recipients | EmailShareFields.Date;
        if ((options.RetainedFields & ~known) != 0) throw new ArgumentOutOfRangeException(nameof(options.RetainedFields));
        if (options.MaxTextChars <= 0 || options.MaxRecipients <= 0 || options.MaxHeaders <= 0 || options.MaxAttachments <= 0 ||
            options.MaxAttachmentBytes <= 0 || options.MaxTotalAttachmentBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Share-copy bounds must be positive.");
        if (options.AttachmentIndexes.Count > options.MaxAttachments) throw new ArgumentException("Too many selected attachments.", nameof(options));
        if (options.RetainedHeaderNames.Count > 64 || options.AddressReplacements.Count > options.MaxRecipients || options.AttachmentNames.Count > options.MaxAttachments)
            throw new ArgumentException("The field-selection policy exceeds its bounds.", nameof(options));
        foreach (string name in options.RetainedHeaderNames) {
            if (string.IsNullOrWhiteSpace(name) || name.Length > 256 || name.Any(character => character <= ' ' || character >= 127 || character == ':'))
                throw new ArgumentException("Retained header names must be valid ASCII field names.", nameof(options));
        }
    }
}
