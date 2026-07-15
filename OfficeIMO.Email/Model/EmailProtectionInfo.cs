namespace OfficeIMO.Email;

/// <summary>Classifies protected message content without performing cryptographic operations.</summary>
public enum EmailProtectionKind {
    /// <summary>No protected-message wrapper was detected.</summary>
    None = 0,
    /// <summary>An opaque S/MIME CMS payload was detected; it can be signed, encrypted, or both.</summary>
    SmimeOpaque = 1,
    /// <summary>A clear-signed S/MIME MIME entity was detected.</summary>
    SmimeClearSigned = 2,
    /// <summary>An OpenPGP/MIME encrypted multipart entity was detected.</summary>
    PgpMimeEncrypted = 3,
    /// <summary>An OpenPGP/MIME clear-signed multipart entity was detected.</summary>
    PgpMimeClearSigned = 4,
    /// <summary>A clear-signed MIME entity with an unrecognized signature protocol was detected.</summary>
    MimeClearSigned = 5,
    /// <summary>An encrypted MIME entity with an unrecognized encryption protocol was detected.</summary>
    MimeEncrypted = 6
}

/// <summary>
/// Describes a protected message and points to the original payload that a cryptographic owner can process.
/// </summary>
public sealed class EmailProtectionInfo {
    /// <summary>Detected protection wrapper.</summary>
    public EmailProtectionKind Kind { get; internal set; }

    /// <summary>Original Outlook message class used for classification, when one exists.</summary>
    public string? MessageClass { get; internal set; }

    /// <summary>
    /// Attachment containing the CMS, signature, or encrypted MIME payload when it can be projected independently.
    /// </summary>
    public EmailAttachment? PayloadAttachment { get; internal set; }

    /// <summary>True when a protected-message wrapper was detected.</summary>
    public bool IsProtected => Kind != EmailProtectionKind.None;
}
