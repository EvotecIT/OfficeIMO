namespace OfficeIMO.Email;

/// <summary>Contains the available body alternatives for an email or Outlook item.</summary>
public sealed class EmailBody {
    /// <summary>Plain-text alternative.</summary>
    public string? Text { get; set; }

    /// <summary>HTML alternative.</summary>
    public string? Html { get; set; }

    /// <summary>Decompressed, byte-preserving RTF source when present.</summary>
    public string? Rtf { get; set; }

    /// <summary>Declared charset for the selected plain-text body.</summary>
    public string? TextCharset { get; set; }

    /// <summary>Declared charset for the selected HTML body.</summary>
    public string? HtmlCharset { get; set; }
}

/// <summary>Represents a file, inline resource, or embedded item attachment.</summary>
public sealed class EmailAttachment {
    private readonly List<MapiProperty> _mapiProperties = new List<MapiProperty>();
    private readonly Dictionary<string, byte[]> _structuredStorageStreams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
    private readonly List<TnefAttribute> _tnefAttributes = new List<TnefAttribute>();
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

    /// <summary>MSG attachment method, such as 1 for by-value or 5 for embedded message.</summary>
    public int? MapiAttachMethod { get; set; }

    /// <summary>Attachment-level MAPI properties.</summary>
    public IList<MapiProperty> MapiProperties => _mapiProperties;

    /// <summary>Relative CFB streams retained for an OLE or custom-storage attachment.</summary>
    public IDictionary<string, byte[]> StructuredStorageStreams => _structuredStorageStreams;

    /// <summary>Ordered raw attachment-level TNEF attributes.</summary>
    public IList<TnefAttribute> TnefAttributes => _tnefAttributes;
}
