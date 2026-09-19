using System.Text.Json.Serialization;

namespace OfficeIMO.Confluence;

/// <summary>Confluence attachment metadata.</summary>
public sealed class ConfluenceAttachment {
    /// <summary>Gets or sets the attachment identifier returned by Confluence.</summary>
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;
    /// <summary>Gets or sets the attachment title used as its file name.</summary>
    [JsonPropertyName("title")]
    public string FileName { get; set; } = string.Empty;
    /// <summary>Gets or sets the reported media type, when supplied.</summary>
    [JsonPropertyName("mediaType")]
    public string? MediaType { get; set; }
    /// <summary>Gets or sets the reported attachment size in bytes.</summary>
    [JsonPropertyName("fileSize")]
    public long FileSize { get; set; }
    /// <summary>Gets or sets the identifier of the page associated with the attachment, when supplied.</summary>
    [JsonPropertyName("pageId")]
    public string? PageId { get; set; }
    /// <summary>Gets or sets the download link supplied by Confluence, when available.</summary>
    [JsonPropertyName("downloadLink")]
    public string? DownloadLink { get; set; }
    /// <summary>Gets or sets the attachment version metadata.</summary>
    [JsonPropertyName("version")]
    public ConfluencePageVersion Version { get; set; } = new ConfluencePageVersion();
}

/// <summary>A cursor-addressable attachment batch.</summary>
public sealed class ConfluenceAttachmentBatch {
    internal ConfluenceAttachmentBatch(IReadOnlyList<ConfluenceAttachment> attachments, string? nextRelativeUri) {
        Attachments = attachments;
        NextRelativeUri = nextRelativeUri;
    }
    /// <summary>Gets the attachments returned in this batch.</summary>
    public IReadOnlyList<ConfluenceAttachment> Attachments { get; }
    /// <summary>Gets the next-page link supplied by Confluence, or <see langword="null"/> at the end.</summary>
    /// <remarks>Despite the property name, an HTTP Link header can supply an absolute URI.</remarks>
    public string? NextRelativeUri { get; }
    /// <summary>Decoded cursor for requesting the next batch, or null when enumeration is complete.</summary>
    public string? NextCursor => ConfluenceCursor.Extract(NextRelativeUri);
}

/// <summary>Attachment data for Confluence's multipart upload endpoint.</summary>
public sealed class ConfluenceAttachmentUpload {
    /// <summary>Gets or sets the required file name sent in the multipart upload.</summary>
    public string FileName { get; set; } = string.Empty;
    /// <summary>Gets or sets the file media type; a blank value is sent as <c>application/octet-stream</c>.</summary>
    public string ContentType { get; set; } = "application/octet-stream";
    /// <summary>Gets or sets the bytes to upload; the request model does not copy the array.</summary>
    public byte[] Content { get; set; } = Array.Empty<byte>();
    /// <summary>Gets or sets an optional attachment comment, omitted when blank.</summary>
    public string? Comment { get; set; }
    /// <summary>Gets or sets whether the upload is marked as a minor edit; defaults to <see langword="true"/>.</summary>
    public bool MinorEdit { get; set; } = true;
}

/// <summary>Streaming attachment data for Confluence's multipart upload endpoint.</summary>
public sealed class ConfluenceAttachmentStreamUpload {
    /// <summary>Gets or sets the required file name sent in the multipart upload.</summary>
    public string FileName { get; set; } = string.Empty;
    /// <summary>Gets or sets the file media type; a blank value is sent as <c>application/octet-stream</c>.</summary>
    public string ContentType { get; set; } = "application/octet-stream";
    /// <summary>Gets or sets the readable stream sent from its current position. Ownership remains with the caller.</summary>
    public Stream Content { get; set; } = Stream.Null;
    /// <summary>Gets or sets an optional attachment comment, omitted when blank.</summary>
    public string? Comment { get; set; }
    /// <summary>Gets or sets whether the upload is marked as a minor edit; defaults to <see langword="true"/>.</summary>
    public bool MinorEdit { get; set; } = true;
}
