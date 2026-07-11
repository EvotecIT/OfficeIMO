namespace OfficeIMO.Email;

/// <summary>Classifies protected message content without performing cryptographic operations.</summary>
public enum EmailProtectionKind {
    /// <summary>No protected-message wrapper was detected.</summary>
    None = 0,
    /// <summary>An opaque S/MIME CMS payload was detected; it can be signed, encrypted, or both.</summary>
    SmimeOpaque = 1,
    /// <summary>A clear-signed S/MIME MIME entity was detected.</summary>
    SmimeClearSigned = 2
}

/// <summary>
/// Describes a protected Outlook message and points to the original payload that a cryptographic owner can process.
/// </summary>
public sealed class EmailProtectionInfo {
    /// <summary>Detected protection wrapper.</summary>
    public EmailProtectionKind Kind { get; internal set; }

    /// <summary>Original Outlook message class used for classification.</summary>
    public string? MessageClass { get; internal set; }

    /// <summary>
    /// Attachment containing the CMS or signed MIME payload. Its content can be handed to MimeKit or another
    /// cryptographic provider when attachment content was retained by the reader.
    /// </summary>
    public EmailAttachment? PayloadAttachment { get; internal set; }

    /// <summary>True when a protected-message wrapper was detected.</summary>
    public bool IsProtected => Kind != EmailProtectionKind.None;
}
