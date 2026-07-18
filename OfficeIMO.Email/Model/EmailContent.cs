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

    /// <summary>Content-ID attached to the selected HTML MIME body part, without angle brackets.</summary>
    public string? HtmlContentId { get; set; }

    /// <summary>Content location attached to the selected HTML MIME body part.</summary>
    public string? HtmlContentLocation { get; set; }

    /// <summary>
    /// True when the selected HTML part is the root of a MIME <c>multipart/related</c> entity. Writers preserve the
    /// related container even when it currently has no resource parts.
    /// </summary>
    public bool IsHtmlRelatedRoot { get; set; }
}

/// <summary>Represents a file, inline resource, or embedded item attachment.</summary>
public sealed class EmailAttachment {
    private readonly List<MapiProperty> _mapiProperties = new List<MapiProperty>();
    private MapiPropertyBag? _mapi;
    private readonly Dictionary<string, string> _contentTypeParameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, byte[]> _structuredStorageStreams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
    private readonly List<TnefAttribute> _tnefAttributes = new List<TnefAttribute>();
    internal bool IsProjectedSemanticContent { get; set; }
    internal bool IsMimeAttachment { get; set; }
    internal bool IsMimeBodyPart { get; set; }
    /// <summary>Attachment filename.</summary>
    public string? FileName { get; set; }

    /// <summary>Declared MIME content type.</summary>
    public string? ContentType { get; set; }

    /// <summary>Declared MIME content-type parameters other than the attachment filename.</summary>
    public IDictionary<string, string> ContentTypeParameters => _contentTypeParameters;

    /// <summary>Content-ID used by inline references.</summary>
    public string? ContentId { get; set; }

    /// <summary>Content location used by inline references.</summary>
    public string? ContentLocation { get; set; }

    /// <summary>True when the source marks the attachment as inline.</summary>
    public bool IsInline { get; set; }

    /// <summary>
    /// True when the source places this part inside the MIME related-resource set. Writers also infer related
    /// membership from exact HTML Content-ID and Content-Location references.
    /// </summary>
    public bool IsMimeRelated { get; set; }

    /// <summary>True when Outlook marks the attachment hidden.</summary>
    public bool IsHidden { get; set; }

    /// <summary>True when the attachment is the picture for an Outlook contact.</summary>
    public bool IsContactPhoto { get; set; }

    /// <summary>RTF rendering position, or -1 when not rendered in the body.</summary>
    public int RenderingPosition { get; set; } = -1;

    /// <summary>Attachment creation timestamp.</summary>
    public DateTimeOffset? CreatedDate { get; set; }

    /// <summary>Attachment modification timestamp.</summary>
    public DateTimeOffset? ModifiedDate { get; set; }

    /// <summary>Linked attachment path for by-reference attachment methods.</summary>
    public string? LinkedPath { get; set; }

    /// <summary>Decoded payload length.</summary>
    public long Length { get; set; }

    /// <summary>Decoded content when requested by reader options.</summary>
    public byte[]? Content { get; set; }

    /// <summary>
    /// Reopenable decoded content used when retaining a byte array is undesirable. Writers prefer
    /// <see cref="Content"/> when both representations are present.
    /// </summary>
    public IEmailContentSource? ContentSource { get; set; }

    /// <summary>Opens decoded attachment content without transferring ownership of the attachment.</summary>
    public Stream OpenContentStream() {
        if (Content != null) return new MemoryStream(Content, writable: false);
        if (ContentSource != null) {
            Stream stream = ContentSource.OpenRead();
            if (stream == null || !stream.CanRead) {
                stream?.Dispose();
                throw new InvalidDataException("The attachment content source did not return a readable stream.");
            }
            return stream;
        }
        return new MemoryStream(Array.Empty<byte>(), writable: false);
    }

    /// <summary>Asynchronously opens decoded attachment content.</summary>
    public async Task<Stream> OpenContentStreamAsync(CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (Content != null) return new MemoryStream(Content, writable: false);
        if (ContentSource != null) {
            Stream stream = await ContentSource.OpenReadAsync(cancellationToken).ConfigureAwait(false);
            if (stream == null || !stream.CanRead) {
                stream?.Dispose();
                throw new InvalidDataException("The attachment content source did not return a readable stream.");
            }
            return stream;
        }
        return new MemoryStream(Array.Empty<byte>(), writable: false);
    }

    /// <summary>Embedded message or Outlook item when the attachment is structured.</summary>
    public EmailDocument? EmbeddedDocument { get; set; }

    /// <summary>MSG attachment method, such as 1 for by-value or 5 for embedded message.</summary>
    public int? MapiAttachMethod { get; set; }

    /// <summary>Attachment-level MAPI properties.</summary>
    public IList<MapiProperty> MapiProperties => _mapiProperties;

    /// <summary>Typed MAPI access backed by the exact <see cref="MapiProperties"/> collection.</summary>
    public MapiPropertyBag Mapi => _mapi ?? (_mapi = new MapiPropertyBag(_mapiProperties));

    /// <summary>Relative CFB streams retained for an OLE, embedded MSG, or custom-storage attachment.</summary>
    public IDictionary<string, byte[]> StructuredStorageStreams => _structuredStorageStreams;

    /// <summary>Ordered raw attachment-level TNEF attributes.</summary>
    public IList<TnefAttribute> TnefAttributes => _tnefAttributes;
}
