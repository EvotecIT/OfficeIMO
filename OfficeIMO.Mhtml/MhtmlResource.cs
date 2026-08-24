using OfficeIMO.Email;

namespace OfficeIMO.Mhtml;

/// <summary>Immutable decoded resource embedded in an MHTML web archive.</summary>
public sealed class MhtmlResource {
    private readonly byte[] _content;

    /// <summary>Creates an embedded MHTML resource snapshot.</summary>
    public MhtmlResource(byte[] content, string? contentType = null, string? contentId = null,
        string? contentLocation = null, string? fileName = null)
        : this(content, contentType, contentId, contentLocation, fileName, takeOwnership: false) {
    }

    private MhtmlResource(byte[] content, string? contentType, string? contentId,
        string? contentLocation, string? fileName, bool takeOwnership) {
        if (content == null) throw new ArgumentNullException(nameof(content));
        if (string.IsNullOrWhiteSpace(contentId) && string.IsNullOrWhiteSpace(contentLocation) &&
            string.IsNullOrWhiteSpace(fileName)) {
            throw new ArgumentException("An MHTML resource requires a Content-ID, Content-Location, or filename.");
        }
        _content = takeOwnership ? content : (byte[])content.Clone();
        ContentType = string.IsNullOrWhiteSpace(contentType) ? "application/octet-stream" : contentType!.Trim();
        ContentId = NormalizeContentId(contentId);
        ContentLocation = string.IsNullOrWhiteSpace(contentLocation) ? null : contentLocation!.Trim();
        FileName = string.IsNullOrWhiteSpace(fileName) ? null : fileName!.Trim();
    }

    /// <summary>Declared MIME content type.</summary>
    public string ContentType { get; }

    /// <summary>Content-ID without angle brackets.</summary>
    public string? ContentId { get; }

    /// <summary>Content location as declared by the archive.</summary>
    public string? ContentLocation { get; }

    /// <summary>Optional resource filename.</summary>
    public string? FileName { get; }

    /// <summary>Decoded content length.</summary>
    public long Length => _content.LongLength;

    /// <summary>Returns an independent copy of the decoded content.</summary>
    public byte[] Content => (byte[])_content.Clone();

    // MHTML is immutable, so the package resolver can safely borrow this snapshot and avoid a
    // public-copy followed by a second constructor copy for every render.
    internal byte[] EncodedContent => _content;

    /// <summary>Opens an independent read-only content stream.</summary>
    public Stream OpenRead() => new MemoryStream(_content, writable: false);

    internal EmailAttachment ToEmailAttachment() => new EmailAttachment {
        FileName = FileName,
        ContentType = ContentType,
        ContentId = ContentId,
        ContentLocation = ContentLocation,
        IsInline = true,
        IsMimeRelated = true,
        Content = _content,
        Length = _content.LongLength
    };

    internal static MhtmlResource FromEmailAttachment(EmailAttachment attachment) {
        if (attachment == null) throw new ArgumentNullException(nameof(attachment));
        if (attachment.Content != null) {
            return new MhtmlResource(
                attachment.Content,
                attachment.ContentType,
                attachment.ContentId,
                attachment.ContentLocation,
                attachment.FileName,
                takeOwnership: true);
        }

        using Stream stream = attachment.OpenContentStream();
        int capacity = attachment.Length is > 0 and <= int.MaxValue ? (int)attachment.Length : 0;
        using var output = capacity == 0 ? new MemoryStream() : new MemoryStream(capacity);
        stream.CopyTo(output);
        return new MhtmlResource(
            output.ToArray(),
            attachment.ContentType,
            attachment.ContentId,
            attachment.ContentLocation,
            attachment.FileName,
            takeOwnership: true);
    }

    private static string? NormalizeContentId(string? contentId) {
        if (string.IsNullOrWhiteSpace(contentId)) return null;
        return contentId!.Trim().Trim('<', '>');
    }
}
