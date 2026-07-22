using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Confluence;

public sealed partial class ConfluenceClient {
    /// <summary>Lists one cursor-addressable batch of attachments for a page.</summary>
    public async Task<ConfluenceAttachmentBatch> GetAttachmentsAsync(string pageId, string? cursor = null, int limit = 50, CancellationToken cancellationToken = default) {
        ValidateId(pageId, nameof(pageId));
        if (limit < 1 || limit > 250) throw new ArgumentOutOfRangeException(nameof(limit), "Limit must be between 1 and 250.");
        string uri = "/wiki/api/v2/pages/" + Encode(pageId) + "/attachments?limit=" + limit.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (!string.IsNullOrWhiteSpace(cursor)) uri += "&cursor=" + Encode(cursor!);
        ConfluenceJsonResponse<ConfluenceCollectionResponse<ConfluenceAttachment>> response = await _transport.SendJsonWithHeadersAsync(HttpMethod.Get, uri, null, ConfluenceRequestSafety.SafeToRetry, ConfluenceJsonSerializerContext.Default.ConfluenceCollectionResponseConfluenceAttachment, cancellationToken).ConfigureAwait(false);
        return new ConfluenceAttachmentBatch(response.Value.Results, ConfluencePagination.Next(response.Value.Links?.Next, response.Headers));
    }

    /// <summary>Downloads an attachment through Confluence's authenticated redirect endpoint.</summary>
    public async Task<byte[]> DownloadAttachmentAsync(string pageId, string attachmentId, CancellationToken cancellationToken = default) {
        ValidateId(pageId, nameof(pageId));
        ValidateId(attachmentId, nameof(attachmentId));
        string uri = "/wiki/rest/api/content/" + Encode(pageId) + "/child/attachment/" + Encode(attachmentId) + "/download";
        using var destination = new MemoryStream();
        await DownloadAttachmentAsync(pageId, attachmentId, destination, cancellationToken).ConfigureAwait(false);
        return destination.ToArray();
    }

    /// <summary>Streams an attachment through Confluence's authenticated redirect endpoint.</summary>
    public Task DownloadAttachmentAsync(string pageId, string attachmentId, Stream destination, CancellationToken cancellationToken = default) {
        ValidateId(pageId, nameof(pageId));
        ValidateId(attachmentId, nameof(attachmentId));
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        string uri = "/wiki/rest/api/content/" + Encode(pageId) + "/child/attachment/" + Encode(attachmentId) + "/download";
        return _transport.SendToStreamAsync(HttpMethod.Get, uri, destination, cancellationToken);
    }

    /// <summary>Creates or versions an attachment through Confluence's multipart v1 endpoint.</summary>
    public async Task<IReadOnlyList<ConfluenceAttachment>> UploadAttachmentAsync(string pageId, ConfluenceAttachmentUpload upload, CancellationToken cancellationToken = default) {
        ValidateId(pageId, nameof(pageId));
        if (upload == null) throw new ArgumentNullException(nameof(upload));
        if (string.IsNullOrWhiteSpace(upload.FileName)) throw new ArgumentException("Attachment file name is required.", nameof(upload));
        if (upload.Content == null) throw new ArgumentException("Attachment content is required.", nameof(upload));
        using var content = new MemoryStream(upload.Content, writable: false);
        return await UploadAttachmentAsync(pageId, new ConfluenceAttachmentStreamUpload {
            FileName = upload.FileName,
            ContentType = upload.ContentType,
            Content = content,
            Comment = upload.Comment,
            MinorEdit = upload.MinorEdit,
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Creates or versions an attachment from a caller-owned stream.</summary>
    public async Task<IReadOnlyList<ConfluenceAttachment>> UploadAttachmentAsync(string pageId, ConfluenceAttachmentStreamUpload upload, CancellationToken cancellationToken = default) {
        ValidateId(pageId, nameof(pageId));
        if (upload == null) throw new ArgumentNullException(nameof(upload));
        if (string.IsNullOrWhiteSpace(upload.FileName)) throw new ArgumentException("Attachment file name is required.", nameof(upload));
        if (upload.Content == null || !upload.Content.CanRead) throw new ArgumentException("A readable attachment content stream is required.", nameof(upload));
        string uri = "/wiki/rest/api/content/" + Encode(pageId) + "/child/attachment";
        ConfluenceHttpResponse response = await _transport.SendMultipartAsync(uri, () => CreateMultipart(upload), cancellationToken).ConfigureAwait(false);
        ConfluenceV1AttachmentResponse? parsed = JsonSerializer.Deserialize(response.Body, ConfluenceJsonSerializerContext.Default.ConfluenceV1AttachmentResponse);
        return parsed == null
            ? Array.Empty<ConfluenceAttachment>()
            : parsed.Results.Select(item => new ConfluenceAttachment {
                Id = item.Id,
                FileName = item.Title,
                MediaType = item.Extensions?.MediaType ?? item.Metadata?.MediaType,
                FileSize = item.Extensions?.FileSize ?? item.FileSize,
                PageId = pageId,
                DownloadLink = item.Links?.Download,
                Version = item.Version ?? new ConfluencePageVersion(),
            }).ToArray();
    }

    private static HttpContent CreateMultipart(ConfluenceAttachmentStreamUpload upload) {
        var multipart = new MultipartFormDataContent();
        var file = new StreamContent(new NonDisposingReadStream(upload.Content));
        file.Headers.ContentType = new MediaTypeHeaderValue(string.IsNullOrWhiteSpace(upload.ContentType) ? "application/octet-stream" : upload.ContentType);
        multipart.Add(file, "file", upload.FileName);
        multipart.Add(new StringContent(upload.MinorEdit ? "true" : "false", Encoding.UTF8, "text/plain"), "minorEdit");
        if (!string.IsNullOrWhiteSpace(upload.Comment)) multipart.Add(new StringContent(upload.Comment!, Encoding.UTF8, "text/plain"), "comment");
        return multipart;
    }

    private sealed class NonDisposingReadStream : Stream {
        private readonly Stream _inner;
        internal NonDisposingReadStream(Stream inner) => _inner = inner;
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => _inner.CanSeek;
        public override bool CanWrite => false;
        public override long Length => _inner.Length;
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        public override void Flush() => _inner.Flush();
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { }
    }

}

internal sealed class ConfluenceV1AttachmentResponse {
    [System.Text.Json.Serialization.JsonPropertyName("results")]
    public List<ConfluenceV1Attachment> Results { get; set; } = new List<ConfluenceV1Attachment>();
}

internal sealed class ConfluenceV1Attachment {
    [System.Text.Json.Serialization.JsonPropertyName("id")]
    public string Id { get; set; } = string.Empty;
    [System.Text.Json.Serialization.JsonPropertyName("title")]
    public string Title { get; set; } = string.Empty;
    [System.Text.Json.Serialization.JsonPropertyName("fileSize")]
    public long FileSize { get; set; }
    [System.Text.Json.Serialization.JsonPropertyName("metadata")]
    public ConfluenceV1AttachmentMetadata? Metadata { get; set; }
    [System.Text.Json.Serialization.JsonPropertyName("extensions")]
    public ConfluenceV1AttachmentExtensions? Extensions { get; set; }
    [System.Text.Json.Serialization.JsonPropertyName("_links")]
    public ConfluenceV1AttachmentLinks? Links { get; set; }
    [System.Text.Json.Serialization.JsonPropertyName("version")]
    public ConfluencePageVersion? Version { get; set; }
}

internal sealed class ConfluenceV1AttachmentMetadata {
    [System.Text.Json.Serialization.JsonPropertyName("mediaType")]
    public string? MediaType { get; set; }
}

internal sealed class ConfluenceV1AttachmentExtensions {
    [System.Text.Json.Serialization.JsonPropertyName("mediaType")]
    public string? MediaType { get; set; }
    [System.Text.Json.Serialization.JsonPropertyName("fileSize")]
    public long FileSize { get; set; }
}

internal sealed class ConfluenceV1AttachmentLinks {
    [System.Text.Json.Serialization.JsonPropertyName("download")]
    public string? Download { get; set; }
}
