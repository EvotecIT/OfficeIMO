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
        CollectionResponse<ConfluenceAttachment> response = await _transport.SendJsonAsync<CollectionResponse<ConfluenceAttachment>>(HttpMethod.Get, uri, null, ConfluenceRequestSafety.SafeToRetry, cancellationToken).ConfigureAwait(false);
        return new ConfluenceAttachmentBatch(response.Results, response.Links?.Next);
    }

    /// <summary>Downloads an attachment through Confluence's authenticated redirect endpoint.</summary>
    public async Task<byte[]> DownloadAttachmentAsync(string pageId, string attachmentId, CancellationToken cancellationToken = default) {
        ValidateId(pageId, nameof(pageId));
        ValidateId(attachmentId, nameof(attachmentId));
        string uri = "/wiki/rest/api/content/" + Encode(pageId) + "/child/attachment/" + Encode(attachmentId) + "/download";
        ConfluenceHttpResponse response = await _transport.SendRawAsync(HttpMethod.Get, uri, cancellationToken).ConfigureAwait(false);
        return response.Body;
    }

    /// <summary>Creates or versions an attachment through Confluence's multipart v1 endpoint.</summary>
    public async Task<IReadOnlyList<ConfluenceAttachment>> UploadAttachmentAsync(string pageId, ConfluenceAttachmentUpload upload, CancellationToken cancellationToken = default) {
        ValidateId(pageId, nameof(pageId));
        if (upload == null) throw new ArgumentNullException(nameof(upload));
        if (string.IsNullOrWhiteSpace(upload.FileName)) throw new ArgumentException("Attachment file name is required.", nameof(upload));
        if (upload.Content == null) throw new ArgumentException("Attachment content is required.", nameof(upload));
        string uri = "/wiki/rest/api/content/" + Encode(pageId) + "/child/attachment";
        ConfluenceHttpResponse response = await _transport.SendMultipartAsync(uri, () => CreateMultipart(upload), cancellationToken).ConfigureAwait(false);
        V1AttachmentResponse? parsed = JsonSerializer.Deserialize<V1AttachmentResponse>(response.Body, ConfluenceHttpTransport.JsonOptions);
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

    private static HttpContent CreateMultipart(ConfluenceAttachmentUpload upload) {
        var multipart = new MultipartFormDataContent();
        var file = new ByteArrayContent(upload.Content);
        file.Headers.ContentType = new MediaTypeHeaderValue(string.IsNullOrWhiteSpace(upload.ContentType) ? "application/octet-stream" : upload.ContentType);
        multipart.Add(file, "file", upload.FileName);
        multipart.Add(new StringContent(upload.MinorEdit ? "true" : "false", Encoding.UTF8, "text/plain"), "minorEdit");
        if (!string.IsNullOrWhiteSpace(upload.Comment)) multipart.Add(new StringContent(upload.Comment!, Encoding.UTF8, "text/plain"), "comment");
        return multipart;
    }

    private sealed class V1AttachmentResponse {
        [System.Text.Json.Serialization.JsonPropertyName("results")]
        public List<V1Attachment> Results { get; set; } = new List<V1Attachment>();
    }

    private sealed class V1Attachment {
        [System.Text.Json.Serialization.JsonPropertyName("id")]
        public string Id { get; set; } = string.Empty;
        [System.Text.Json.Serialization.JsonPropertyName("title")]
        public string Title { get; set; } = string.Empty;
        [System.Text.Json.Serialization.JsonPropertyName("fileSize")]
        public long FileSize { get; set; }
        [System.Text.Json.Serialization.JsonPropertyName("metadata")]
        public V1AttachmentMetadata? Metadata { get; set; }
        [System.Text.Json.Serialization.JsonPropertyName("extensions")]
        public V1AttachmentExtensions? Extensions { get; set; }
        [System.Text.Json.Serialization.JsonPropertyName("_links")]
        public V1AttachmentLinks? Links { get; set; }
        [System.Text.Json.Serialization.JsonPropertyName("version")]
        public ConfluencePageVersion? Version { get; set; }
    }

    private sealed class V1AttachmentMetadata {
        [System.Text.Json.Serialization.JsonPropertyName("mediaType")]
        public string? MediaType { get; set; }
    }

    private sealed class V1AttachmentExtensions {
        [System.Text.Json.Serialization.JsonPropertyName("mediaType")]
        public string? MediaType { get; set; }
        [System.Text.Json.Serialization.JsonPropertyName("fileSize")]
        public long FileSize { get; set; }
    }

    private sealed class V1AttachmentLinks {
        [System.Text.Json.Serialization.JsonPropertyName("download")]
        public string? Download { get; set; }
    }
}
