namespace OfficeIMO.Email;

/// <summary>Envelope fields explicitly retained in a share artifact.</summary>
[Flags]
public enum EmailShareFields {
    /// <summary>Retain no original envelope fields.</summary>
    None = 0,
    /// <summary>Retain the original subject unless replaced.</summary>
    Subject = 1,
    /// <summary>Retain the represented author, subject to address replacements.</summary>
    Author = 2,
    /// <summary>Retain To and Cc recipients, subject to address replacements.</summary>
    Recipients = 4,
    /// <summary>Retain the declared date.</summary>
    Date = 8
}

/// <summary>Explicit field selection for an independent plain-text artifact; this is not complete anonymization.</summary>
public sealed class EmailShareCopyOptions {
    /// <summary>Original envelope fields to retain. None are retained by default.</summary>
    public EmailShareFields RetainedFields { get; set; }
    /// <summary>Replacement subject, including an empty value. Null uses the subject selection policy.</summary>
    public string? ReplacementSubject { get; set; }
    /// <summary>Replacement plain body, including an empty value. Null uses the original plain-text body.</summary>
    public string? ReplacementBodyText { get; set; }
    /// <summary>Replace selected address values, or omit them when the replacement is null. Matching ignores case.</summary>
    public IDictionary<string, string?> AddressReplacements { get; } = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
    /// <summary>Retain address display names. Raw spelling and source-specific recipient properties are never copied.</summary>
    public bool KeepDisplayNames { get; set; }
    /// <summary>Explicit extra header names to retain; envelope, MIME and integrity headers remain owned by the new artifact.</summary>
    public ICollection<string> RetainedHeaderNames { get; } = new List<string>();
    /// <summary>Zero-based ordinary attachment indexes to copy. Payload content is retained unchanged and may contain private data.</summary>
    public ICollection<int> AttachmentIndexes { get; } = new List<int>();
    /// <summary>Retain selected original attachment names unless overridden. Generic names are used by default.</summary>
    public bool KeepAttachmentNames { get; set; }
    /// <summary>Replacement names for selected attachments, keyed by original index.</summary>
    public IDictionary<int, string> AttachmentNames { get; } = new Dictionary<int, string>();
    /// <summary>Maximum source or replacement text field length, before any artifact is created.</summary>
    public int MaxTextChars { get; set; } = 2 * 1024 * 1024;
    /// <summary>Maximum original recipients considered.</summary>
    public int MaxRecipients { get; set; } = 1000;
    /// <summary>Maximum original headers considered.</summary>
    public int MaxHeaders { get; set; } = 10_000;
    /// <summary>Maximum selected attachment count.</summary>
    public int MaxAttachments { get; set; } = 100;
    /// <summary>Maximum decoded bytes per selected attachment.</summary>
    public long MaxAttachmentBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum decoded bytes across selected attachments.</summary>
    public long MaxTotalAttachmentBytes { get; set; } = 256L * 1024 * 1024;

    /// <summary>Creates an independent policy snapshot, including selection collections.</summary>
    public EmailShareCopyOptions Clone() {
        var clone = new EmailShareCopyOptions { RetainedFields = RetainedFields, ReplacementSubject = ReplacementSubject,
            ReplacementBodyText = ReplacementBodyText, KeepDisplayNames = KeepDisplayNames, KeepAttachmentNames = KeepAttachmentNames,
            MaxTextChars = MaxTextChars, MaxRecipients = MaxRecipients, MaxHeaders = MaxHeaders, MaxAttachments = MaxAttachments,
            MaxAttachmentBytes = MaxAttachmentBytes, MaxTotalAttachmentBytes = MaxTotalAttachmentBytes };
        foreach (var entry in AddressReplacements) clone.AddressReplacements.Add(entry.Key, entry.Value);
        foreach (string name in RetainedHeaderNames) clone.RetainedHeaderNames.Add(name);
        foreach (int index in AttachmentIndexes) clone.AttachmentIndexes.Add(index);
        foreach (var entry in AttachmentNames) clone.AttachmentNames.Add(entry.Key, entry.Value);
        return clone;
    }
}

/// <summary>Field handling evidence without original field values.</summary>
public sealed class EmailShareFieldChange {
    internal EmailShareFieldChange(string path, string action) { Path = path; Action = action; }
    /// <summary>Model field path, or a zero-based recipient/attachment index.</summary>
    public string Path { get; }
    /// <summary>Copied, replaced or omitted handling.</summary>
    public string Action { get; }
}

/// <summary>An independent share artifact and evidence about omissions and integrity effects.</summary>
public sealed class EmailShareCopyResult {
    internal EmailShareCopyResult(EmailDocument document, IReadOnlyList<EmailShareFieldChange> changes, IReadOnlyList<EmailDiagnostic> diagnostics) {
        Document = document; Changes = changes; Diagnostics = diagnostics;
    }
    /// <summary>New plain-text EML model with no inherited raw source, MAPI/TNEF metadata, Bcc or protected wrapper.</summary>
    public EmailDocument Document { get; }
    /// <summary>Selected field handling, without original values.</summary>
    public IReadOnlyList<EmailShareFieldChange> Changes { get; }
    /// <summary>Body availability, retained opaque payloads and removed integrity evidence.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
}
