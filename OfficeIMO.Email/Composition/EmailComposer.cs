namespace OfficeIMO.Email;

/// <summary>Creates new reply and forward drafts over the format-neutral email model.</summary>
public static class EmailComposer {
    /// <summary>Replies to Reply-To recipients, or From when Reply-To is absent. Own addresses are excluded.</summary>
    public static EmailCompositionResult Reply(EmailDocument original, EmailAddress from, string text,
        EmailCompositionOptions? options = null) => Compose(original, from, text, false, false, options);

    /// <summary>Replies to the reply target and original To/Cc recipients, excluding own addresses and duplicates.</summary>
    public static EmailCompositionResult ReplyAll(EmailDocument original, EmailAddress from, string text,
        EmailCompositionOptions? options = null) => Compose(original, from, text, true, false, options);

    /// <summary>Creates a quoted forward with no recipients or attachments. Add intended recipients and attachments explicitly.</summary>
    public static EmailCompositionResult Forward(EmailDocument original, EmailAddress from, string text,
        EmailCompositionOptions? options = null) => Compose(original, from, text, false, true, options);

    private static EmailCompositionResult Compose(EmailDocument original, EmailAddress from, string text,
        bool replyAll, bool forward, EmailCompositionOptions? options) {
        if (original == null) throw new ArgumentNullException(nameof(original));
        if (from == null) throw new ArgumentNullException(nameof(from));
        if (text == null) throw new ArgumentNullException(nameof(text));
        var effective = options ?? new EmailCompositionOptions();
        if (effective.MaxQuoteChars < 0) throw new ArgumentOutOfRangeException(nameof(options), "MaxQuoteChars must be non-negative.");
        if (effective.MaxReferences < 1 || effective.MaxReferences > 1000) throw new ArgumentOutOfRangeException(nameof(options), "MaxReferences must be from 1 to 1000.");
        string? sender = Smtp(from);
        if (sender == null) throw new ArgumentException("A usable SMTP sender is required.", nameof(from));
        var diagnostics = new List<EmailDiagnostic>();
        string prefix = forward ? "Fwd:" : "Re:";
        string subject = original.Subject ?? string.Empty;
        var draft = new EmailDocument {
            Format = EmailFileFormat.Eml, From = CopyAddress(from, sender), Date = effective.Date ?? DateTimeOffset.UtcNow,
            Subject = subject.TrimStart().StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? subject : prefix + " " + subject
        };
        draft.MessageMetadata.IsDraft = true;
        draft.Body.TextCharset = "utf-8";
        draft.Body.Text = text;
        if (!forward) {
            var own = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { sender };
            foreach (string address in effective.OwnAddresses) {
                string? normalized = Smtp(new EmailAddress(address));
                if (normalized == null) throw new ArgumentException("OwnAddresses must contain usable SMTP addresses.", nameof(options));
                own.Add(normalized);
            }
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            EmailRecipient[] replyTargets = original.Recipients.Where(item => item.Kind == EmailRecipientKind.ReplyTo).ToArray();
            if (replyTargets.Length == 0) Add(original.From, EmailRecipientKind.To);
            else foreach (EmailRecipient target in replyTargets) Add(target.Address, EmailRecipientKind.To);
            if (replyAll) {
                foreach (EmailRecipient recipient in original.Recipients.Where(item => item.Kind == EmailRecipientKind.To)) Add(recipient.Address, EmailRecipientKind.To);
                foreach (EmailRecipient recipient in original.Recipients.Where(item => item.Kind == EmailRecipientKind.Cc)) Add(recipient.Address, EmailRecipientKind.Cc);
            }
            if (draft.Recipients.Count == 0) diagnostics.Add(new EmailDiagnostic("EMAIL_COMPOSITION_NO_RECIPIENTS", "No external SMTP reply recipients were resolved."));
            ApplyThreading(original, draft, effective.MaxReferences, diagnostics);

            void Add(EmailAddress? address, EmailRecipientKind kind) {
                string? normalized = Smtp(address);
                if (normalized == null) {
                    diagnostics.Add(new EmailDiagnostic("EMAIL_COMPOSITION_UNRESOLVED_ADDRESS", "A reply address is missing or requires external address resolution."));
                } else if (!own.Contains(normalized) && seen.Add(normalized)) draft.Recipients.Add(new EmailRecipient(kind, CopyAddress(address!, normalized)));
            }
        }
        if (effective.QuoteOriginal) {
            string? quote = original.Body.Text;
            if (quote == null) diagnostics.Add(new EmailDiagnostic("EMAIL_COMPOSITION_PLAIN_BODY_UNAVAILABLE", "No plain-text body is available for quotation; use the HTML bridge or supply an explicit projection."));
            else {
                if (quote.Length > effective.MaxQuoteChars) {
                    int end = effective.MaxQuoteChars;
                    if (end > 0 && char.IsHighSurrogate(quote[end - 1]) && char.IsLowSurrogate(quote[end])) end--;
                    quote = quote.Substring(0, end);
                    diagnostics.Add(new EmailDiagnostic("EMAIL_COMPOSITION_QUOTE_TRUNCATED", "The original quotation reached MaxQuoteChars."));
                }
                draft.Body.Text += "\r\n\r\n" + (forward ? "---------- Forwarded message ----------\r\n" : string.Empty) +
                    string.Join("\r\n", quote.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').Select(line => "> " + line));
            }
        }
        return new EmailCompositionResult(draft, diagnostics.AsReadOnly());
    }

    private static EmailAddress CopyAddress(EmailAddress address, string smtp) => new EmailAddress(smtp, address.DisplayName) { AddressType = "SMTP" };

    private static string? Smtp(EmailAddress? address) => EmailSmtpAddress.Normalize(address);

    private static void ApplyThreading(EmailDocument original, EmailDocument draft, int maximum, List<EmailDiagnostic> diagnostics) {
        string? parent = MessageId(original.MessageId ?? original.Headers.FirstOrDefault(item => item.Name.Equals("Message-ID", StringComparison.OrdinalIgnoreCase))?.Value);
        draft.MessageMetadata.InReplyToId = parent;
        string? references = original.MessageMetadata.InternetReferences ?? original.Headers.FirstOrDefault(item => item.Name.Equals("References", StringComparison.OrdinalIgnoreCase))?.Value;
        if (string.IsNullOrWhiteSpace(references)) {
            references = MessageId(original.MessageMetadata.InReplyToId ?? original.Headers.FirstOrDefault(item => item.Name.Equals("In-Reply-To", StringComparison.OrdinalIgnoreCase))?.Value);
        }
        var ids = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        if (references != null) {
            int length = Math.Min(references.Length, 64 * 1024);
            int position = 0;
            while (position < length) {
                int start = references.IndexOf('<', position, length - position);
                if (start < 0) break;
                int end = references.IndexOf('>', start + 1, length - start - 1);
                if (end < 0) break;
                string? id = MessageId(references.Substring(start, end - start + 1));
                if (id != null && seen.Add(id)) ids.Add(id);
                position = end + 1;
            }
            if (length < references.Length) diagnostics.Add(new EmailDiagnostic("EMAIL_COMPOSITION_REFERENCES_TRUNCATED", "References input exceeded 64 KiB."));
        }
        if (parent != null) { ids.Remove(parent); ids.Add(parent); }
        if (ids.Count > maximum) {
            ids.RemoveRange(0, ids.Count - maximum);
            diagnostics.Add(new EmailDiagnostic("EMAIL_COMPOSITION_REFERENCES_TRUNCATED", "References exceeded MaxReferences; the most recent identifiers were retained."));
        }
        draft.MessageMetadata.InternetReferences = ids.Count == 0 ? null : string.Join(" ", ids);
    }

    private static string? MessageId(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string id = value!.Trim().Trim('<', '>');
        return id.Length > 0 && id.Length <= 998 && id.Contains("@") && id.All(character => character >= '!' && character <= '~' && character != '<' && character != '>')
            ? "<" + id + ">" : null;
    }
}
