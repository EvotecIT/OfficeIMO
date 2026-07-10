namespace OfficeIMO.Email;

/// <summary>Contains the available body alternatives for an email or Outlook item.</summary>
public sealed class EmailBody {
    /// <summary>Plain-text alternative.</summary>
    public string? Text { get; set; }

    /// <summary>HTML alternative.</summary>
    public string? Html { get; set; }

    /// <summary>Decompressed RTF alternative when present.</summary>
    public string? Rtf { get; set; }

    /// <summary>Declared charset for the selected plain-text body.</summary>
    public string? TextCharset { get; set; }

    /// <summary>Declared charset for the selected HTML body.</summary>
    public string? HtmlCharset { get; set; }
}

/// <summary>Represents a file, inline resource, or embedded item attachment.</summary>
public sealed class EmailAttachment {
    /// <summary>Attachment filename.</summary>
    public string? FileName { get; set; }

    /// <summary>Declared MIME content type.</summary>
    public string? ContentType { get; set; }

    /// <summary>Content-ID used by inline references.</summary>
    public string? ContentId { get; set; }

    /// <summary>Content location used by inline references.</summary>
    public string? ContentLocation { get; set; }

    /// <summary>True when the source marks the attachment as inline.</summary>
    public bool IsInline { get; set; }

    /// <summary>Decoded payload length.</summary>
    public long Length { get; set; }

    /// <summary>Decoded content when requested by reader options.</summary>
    public byte[]? Content { get; set; }

    /// <summary>Embedded message or Outlook item when the attachment is structured.</summary>
    public EmailDocument? EmbeddedDocument { get; set; }
}
