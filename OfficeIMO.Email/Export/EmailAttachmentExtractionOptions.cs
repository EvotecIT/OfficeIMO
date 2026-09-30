namespace OfficeIMO.Email;

/// <summary>Bounds attachment export without following linked attachment paths.</summary>
public sealed class EmailAttachmentExtractionOptions {
    /// <summary>Creates an export policy. Embedded messages are exported as EML; nested attachments are optional.</summary>
    public EmailAttachmentExtractionOptions(int maxAttachments = 1000, long maxAttachmentBytes = 64L * 1024 * 1024,
        long maxTotalBytes = 256L * 1024 * 1024, int maxDepth = 4, bool recurseEmbeddedMessages = false,
        bool includeInline = true, bool includeHidden = false) {
        if (maxAttachments <= 0) throw new ArgumentOutOfRangeException(nameof(maxAttachments));
        if (maxAttachmentBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxAttachmentBytes));
        if (maxTotalBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxTotalBytes));
        if (maxDepth < 0 || maxDepth > 32) throw new ArgumentOutOfRangeException(nameof(maxDepth));
        MaxAttachments = maxAttachments; MaxAttachmentBytes = maxAttachmentBytes; MaxTotalBytes = maxTotalBytes;
        MaxDepth = maxDepth; RecurseEmbeddedMessages = recurseEmbeddedMessages;
        IncludeInline = includeInline; IncludeHidden = includeHidden;
    }
    /// <summary>Maximum operation-wide attachment visits, including skipped entries and embedded-message serialization.</summary>
    public int MaxAttachments { get; }
    /// <summary>Maximum decoded or serialized bytes for one exported attachment.</summary>
    public long MaxAttachmentBytes { get; }
    /// <summary>Maximum bytes consumed across attachments, including failed attempts.</summary>
    public long MaxTotalBytes { get; }
    /// <summary>Maximum embedded-message nesting depth; direct attachments have depth zero.</summary>
    public int MaxDepth { get; }
    /// <summary>Whether to visit attachments inside embedded messages as well as exporting their EML.</summary>
    public bool RecurseEmbeddedMessages { get; }
    /// <summary>Whether to export inline resources.</summary>
    public bool IncludeInline { get; }
    /// <summary>Whether to export attachments marked hidden.</summary>
    public bool IncludeHidden { get; }
}

/// <summary>Outcome for one logical attachment path in the source model.</summary>
public sealed class EmailAttachmentExtractionEntry {
    internal EmailAttachmentExtractionEntry(string sourcePath, string? originalName, string? outputPath,
        long bytesWritten, string? sha256, IReadOnlyList<EmailDiagnostic> diagnostics) {
        SourcePath = sourcePath; OriginalName = originalName; OutputPath = outputPath;
        BytesWritten = bytesWritten; Sha256 = sha256; Diagnostics = diagnostics;
    }
    /// <summary>Stable zero-based attachment indexes, separated by slashes for embedded items.</summary>
    public string SourcePath { get; }
    /// <summary>Untrusted filename declared by the source.</summary>
    public string? OriginalName { get; }
    /// <summary>Committed destination, or null when skipped or failed.</summary>
    public string? OutputPath { get; }
    /// <summary>Bytes in the committed file.</summary>
    public long BytesWritten { get; }
    /// <summary>SHA-256 of the committed decoded attachment or generated embedded EML.</summary>
    public string? Sha256 { get; }
    /// <summary>Skipped content, write failures, or embedded serialization diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
}

/// <summary>Hash and provenance manifest returned by attachment extraction.</summary>
public sealed class EmailAttachmentExtractionResult {
    internal EmailAttachmentExtractionResult(IReadOnlyList<EmailAttachmentExtractionEntry> entries, long bytesRead,
        bool truncated, IReadOnlyList<EmailDiagnostic> diagnostics) {
        Entries = entries; BytesRead = bytesRead; Truncated = truncated; Diagnostics = diagnostics;
    }
    /// <summary>Ordered outcomes, including failures and skipped attachments.</summary>
    public IReadOnlyList<EmailAttachmentExtractionEntry> Entries { get; }
    /// <summary>Bytes consumed or generated during this operation, including rejected payloads.</summary>
    public long BytesRead { get; }
    /// <summary>Whether a count, nesting, or aggregate resource limit stopped traversal.</summary>
    public bool Truncated { get; }
    /// <summary>Operation-level bounds and traversal diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
}
