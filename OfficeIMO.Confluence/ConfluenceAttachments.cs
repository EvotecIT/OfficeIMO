using System.Text.Json.Serialization;

namespace OfficeIMO.Confluence;

/// <summary>Confluence attachment metadata.</summary>
public sealed class ConfluenceAttachment {
    [JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;
    [JsonPropertyName("title")]
    public string FileName { get; set; } = string.Empty;
    [JsonPropertyName("mediaType")]
    public string? MediaType { get; set; }
    [JsonPropertyName("fileSize")]
    public long FileSize { get; set; }
    [JsonPropertyName("pageId")]
    public string? PageId { get; set; }
    [JsonPropertyName("downloadLink")]
    public string? DownloadLink { get; set; }
    [JsonPropertyName("version")]
    public ConfluencePageVersion Version { get; set; } = new ConfluencePageVersion();
}

/// <summary>A cursor-addressable attachment batch.</summary>
public sealed class ConfluenceAttachmentBatch {
    internal ConfluenceAttachmentBatch(IReadOnlyList<ConfluenceAttachment> attachments, string? nextRelativeUri) {
        Attachments = attachments;
        NextRelativeUri = nextRelativeUri;
    }
    public IReadOnlyList<ConfluenceAttachment> Attachments { get; }
    public string? NextRelativeUri { get; }
    /// <summary>Decoded cursor for requesting the next batch, or null when enumeration is complete.</summary>
    public string? NextCursor => ConfluenceCursor.Extract(NextRelativeUri);
}

/// <summary>Attachment data for Confluence's multipart upload endpoint.</summary>
public sealed class ConfluenceAttachmentUpload {
    public string FileName { get; set; } = string.Empty;
    public string ContentType { get; set; } = "application/octet-stream";
    public byte[] Content { get; set; } = Array.Empty<byte>();
    public string? Comment { get; set; }
    public bool MinorEdit { get; set; } = true;
}
