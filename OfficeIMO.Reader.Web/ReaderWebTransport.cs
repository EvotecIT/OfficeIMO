namespace OfficeIMO.Reader.Web;

internal static class ReaderWebTransport {
    internal const string CapabilityId = "officeimo.reader.web";
    internal const int MaximumInitialBufferCapacity = 64 * 1024;

    internal static async Task<ReaderWebDownload> DownloadAsync(
        HttpClient httpClient,
        Uri uri,
        string? sourceName,
        long maxResponseBytes,
        ReaderWebOptions options,
        CancellationToken cancellationToken) {
        ReaderWebUriPolicy.Validate(uri, options);
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(options.RequestTimeout);
        try {
            using var request = new HttpRequestMessage(HttpMethod.Get, uri);
            using HttpResponseMessage response = await httpClient.SendAsync(
                request,
                HttpCompletionOption.ResponseHeadersRead,
                timeout.Token).ConfigureAwait(false);
            Uri finalUri = response.RequestMessage?.RequestUri ?? uri;
            ReaderWebUriPolicy.Validate(finalUri, options);
            response.EnsureSuccessStatusCode();

            long? declaredLength = response.Content.Headers.ContentLength;
            if (declaredLength.HasValue && declaredLength.Value > maxResponseBytes) {
                throw new IOException(
                    "HTTP response exceeds the effective web input byte limit (" +
                    declaredLength.Value.ToString(CultureInfo.InvariantCulture) + " > " +
                    maxResponseBytes.ToString(CultureInfo.InvariantCulture) + ").");
            }

            byte[] bytes = await ReadBoundedContentAsync(
                response.Content,
                declaredLength,
                maxResponseBytes,
                timeout.Token).ConfigureAwait(false);
            string logicalSourceName = ResolveSourceName(
                sourceName,
                finalUri,
                response.Content.Headers.ContentDisposition);
            return new ReaderWebDownload(
                bytes,
                logicalSourceName,
                uri,
                finalUri,
                (int)response.StatusCode,
                response.Content.Headers.ContentType?.ToString(),
                response.Content.Headers.LastModified?.UtcDateTime);
        } catch (OperationCanceledException exception)
            when (!cancellationToken.IsCancellationRequested && timeout.IsCancellationRequested) {
            throw new TimeoutException("Reader Web request exceeded RequestTimeout.", exception);
        }
    }

    private static async Task<byte[]> ReadBoundedContentAsync(
        HttpContent content,
        long? declaredLength,
        long maxResponseBytes,
        CancellationToken cancellationToken) {
        int capacity = GetInitialBufferCapacity(declaredLength);
        using var output = capacity > 0 ? new MemoryStream(capacity) : new MemoryStream();
        using Stream input = await content.ReadAsStreamAsync().ConfigureAwait(false);
        byte[] buffer = new byte[64 * 1024];
        long total = 0;
        while (true) {
            int read = await input.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            if (total > maxResponseBytes - read) {
                throw new IOException(
                    "HTTP response exceeds the effective web input byte limit (" +
                    (total + read).ToString(CultureInfo.InvariantCulture) + " > " +
                    maxResponseBytes.ToString(CultureInfo.InvariantCulture) + ").");
            }
            output.Write(buffer, 0, read);
            total += read;
        }
        return output.ToArray();
    }

    internal static int GetInitialBufferCapacity(long? declaredLength) {
        if (!declaredLength.HasValue || declaredLength.Value <= 0) return 0;
        return (int)Math.Min(declaredLength.Value, MaximumInitialBufferCapacity);
    }

    private static string ResolveSourceName(
        string? sourceName,
        Uri finalUri,
        ContentDispositionHeaderValue? contentDisposition) {
        if (sourceName != null) {
            return sourceName;
        }

        string? candidate = contentDisposition?.FileNameStar ?? contentDisposition?.FileName;
        if (string.IsNullOrWhiteSpace(candidate)) {
            try {
                candidate = Uri.UnescapeDataString(finalUri.AbsolutePath);
            } catch (UriFormatException) {
                candidate = finalUri.AbsolutePath;
            }
        }
        string derived = SanitizeDerivedSourceName(candidate);
        return derived.Length == 0 ? "download" : derived;
    }

    private static string SanitizeDerivedSourceName(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        string normalized = value!.Trim().Trim('"').Replace('\\', '/');
        int separator = normalized.LastIndexOf('/');
        if (separator >= 0) normalized = normalized.Substring(separator + 1);
        var safe = new StringBuilder(Math.Min(normalized.Length, 512));
        for (int index = 0; index < normalized.Length; index++) {
            char character = normalized[index];
            if (!char.IsControl(character)) safe.Append(character);
        }
        string result = safe.ToString().Trim();
        if (result == "." || result == "..") return string.Empty;
        if (result.Length > 512) result = result.Substring(result.Length - 512);
        return result;
    }

    internal static string? NormalizeExplicitSourceName(string? sourceName) {
        if (sourceName == null) return null;
        string normalized = sourceName.Trim();
        if (normalized.Length == 0) {
            throw new ArgumentException("Source name cannot be empty when supplied.", nameof(sourceName));
        }
        if (normalized.Length > 2048) {
            throw new ArgumentException("Source name cannot exceed 2,048 characters.", nameof(sourceName));
        }
        for (int index = 0; index < normalized.Length; index++) {
            if (char.IsControl(normalized[index])) {
                throw new ArgumentException("Source name cannot contain control characters.", nameof(sourceName));
            }
        }
        return normalized;
    }
}

internal sealed class ReaderWebDownload {
    internal ReaderWebDownload(
        byte[] bytes,
        string sourceName,
        Uri requestUri,
        Uri responseUri,
        int statusCode,
        string? contentType,
        DateTime? lastModifiedUtc) {
        Bytes = bytes;
        SourceName = sourceName;
        RequestUri = requestUri;
        ResponseUri = responseUri;
        StatusCode = statusCode;
        ContentType = contentType;
        LastModifiedUtc = lastModifiedUtc;
    }

    internal byte[] Bytes { get; }
    internal string SourceName { get; }
    internal Uri RequestUri { get; }
    internal Uri ResponseUri { get; }
    internal int StatusCode { get; }
    internal string? ContentType { get; }
    internal DateTime? LastModifiedUtc { get; }

    internal void ApplyTransportMetadata(OfficeDocumentReadResult result, ReaderWebOptions options) {
        result.CapabilitiesUsed = result.CapabilitiesUsed
            .Concat(new[] { ReaderWebTransport.CapabilityId })
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        result.Source.LengthBytes = Bytes.LongLength;
        if (!result.Source.LastWriteUtc.HasValue && LastModifiedUtc.HasValue) {
            result.Source.LastWriteUtc = LastModifiedUtc;
        }
        var metadata = new List<OfficeDocumentMetadataEntry> {
            Metadata("reader-web-request-uri", "RequestUri", FormatUri(RequestUri, options.IncludeQueryInMetadata), "uri"),
            Metadata("reader-web-response-uri", "ResponseUri", FormatUri(ResponseUri, options.IncludeQueryInMetadata), "uri"),
            Metadata("reader-web-status-code", "StatusCode", StatusCode.ToString(CultureInfo.InvariantCulture), "number"),
            Metadata("reader-web-length", "LengthBytes", Bytes.LongLength.ToString(CultureInfo.InvariantCulture), "number")
        };
        if (!string.IsNullOrWhiteSpace(ContentType)) {
            metadata.Add(Metadata("reader-web-content-type", "ContentType", ContentType!, "string"));
        }
        if (LastModifiedUtc.HasValue) {
            metadata.Add(Metadata(
                "reader-web-last-modified",
                "LastModifiedUtc",
                LastModifiedUtc.Value.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture),
                "timestamp"));
        }
        result.Metadata = result.Metadata.Concat(metadata).ToArray();
    }

    private static OfficeDocumentMetadataEntry Metadata(string id, string name, string value, string valueType) {
        return new OfficeDocumentMetadataEntry {
            Id = id,
            Category = "web.transport",
            Name = name,
            Value = value,
            ValueType = valueType
        };
    }

    private static string FormatUri(Uri uri, bool includeQuery) {
        var builder = new UriBuilder(uri) { Fragment = string.Empty };
        if (!includeQuery) builder.Query = string.Empty;
        return builder.Uri.AbsoluteUri;
    }
}
