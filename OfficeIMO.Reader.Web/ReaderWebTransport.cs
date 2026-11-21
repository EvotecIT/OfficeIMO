using OfficeIMO.Core.Internal;

namespace OfficeIMO.Reader.Web;

internal static class ReaderWebTransport {
    internal const string CapabilityId = "officeimo.reader.web";
    internal const int MaximumInitialBufferCapacity = 64 * 1024;
    internal const int MaximumRedirects = 10;

    /// <summary>Downloads one response; <paramref name="ownedTransport"/> marks the client from the address policy.</summary>
    internal static async Task<ReaderWebDownload> DownloadAsync(
        HttpClient httpClient,
        bool ownedTransport,
        Uri uri,
        string? sourceName,
        Func<string, long> resolveMaxResponseBytes,
        ReaderWebOptions options,
        CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(resolveMaxResponseBytes);
        ReaderWebUriPolicy.Validate(uri, options, screenHostNames: !ownedTransport);
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(options.RequestTimeout);
        try {
            (HttpResponseMessage response, Uri logicalUri) = await SendWithRedirectsAsync(
                httpClient,
                ownedTransport,
                uri,
                options,
                timeout.Token).ConfigureAwait(false);
            using (response) {
                // A caller-supplied client may follow redirects itself; the policy client never does.
                Uri finalUri = ownedTransport ? logicalUri : response.RequestMessage?.RequestUri ?? logicalUri;
                ReaderWebUriPolicy.Validate(finalUri, options);
                response.EnsureSuccessStatusCode();

                string logicalSourceName = ResolveSourceName(
                    sourceName,
                    finalUri,
                    response.Content.Headers.ContentDisposition);
                long maxResponseBytes = resolveMaxResponseBytes(logicalSourceName);
                long? declaredLength = response.Content.Headers.ContentLength;
                if (declaredLength.HasValue && declaredLength.Value > maxResponseBytes) {
                    throw new IOException(
                        "HTTP response exceeds the effective web input byte limit (" +
                        declaredLength.Value.ToString(CultureInfo.InvariantCulture) + " > " +
                        maxResponseBytes.ToString(CultureInfo.InvariantCulture) + ").");
                }

                MemoryStream content = await ReadBoundedContentAsync(
                    response.Content,
                    declaredLength,
                    maxResponseBytes,
                    timeout.Token).ConfigureAwait(false);
                try {
                    return new ReaderWebDownload(
                        content,
                        logicalSourceName,
                        uri,
                        finalUri,
                        (int)response.StatusCode,
                        response.Content.Headers.ContentType?.ToString(),
                        response.Content.Headers.LastModified?.UtcDateTime);
                } catch {
                    content.Dispose();
                    throw;
                }
            }
        } catch (OperationCanceledException exception)
            when (!cancellationToken.IsCancellationRequested && timeout.IsCancellationRequested) {
            throw new TimeoutException("Reader Web request exceeded RequestTimeout.", exception);
        }
    }

    private static async Task<(HttpResponseMessage Response, Uri LogicalUri)> SendWithRedirectsAsync(
        HttpClient httpClient,
        bool ownedTransport,
        Uri uri,
        ReaderWebOptions options,
        CancellationToken cancellationToken) {
        Uri current = uri;
        for (int redirectCount = 0; ; redirectCount++) {
            HttpResponseMessage response = await SendOneAsync(
                httpClient,
                ownedTransport,
                current,
                options,
                cancellationToken).ConfigureAwait(false);
            if (!IsRedirect(response.StatusCode)) {
                return (response, current);
            }

            Uri? location = response.Headers.Location;
            response.Dispose();
            if (redirectCount >= MaximumRedirects) {
                throw new HttpRequestException(
                    "Reader Web exceeded the maximum of " +
                    MaximumRedirects.ToString(CultureInfo.InvariantCulture) + " redirects.");
            }
            if (location == null) {
                throw new HttpRequestException("Reader Web received a redirect without a Location header.");
            }

            Uri next = location.IsAbsoluteUri ? location : new Uri(current, location);
            if (string.Equals(current.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(next.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) {
                throw new ReaderWebPolicyException(next, "Reader Web rejected a redirect from HTTPS to HTTP.");
            }
            ReaderWebUriPolicy.Validate(next, options, screenHostNames: !ownedTransport);
            if (!ownedTransport && HasDefaultCredentials(httpClient) && !IsSameOrigin(current, next)) {
                // HttpClient adds its default headers to every request, so they would reach the other origin.
                throw new ReaderWebPolicyException(next, "Reader Web rejected a cross-origin redirect because the HTTP client sends default credentials.");
            }
            current = next;
        }
    }

    private static async Task<HttpResponseMessage> SendOneAsync(
        HttpClient httpClient,
        bool ownedTransport,
        Uri uri,
        ReaderWebOptions options,
        CancellationToken cancellationToken) {
        HttpRequestMessage? request = null;
        try {
            request = ownedTransport
                ? await OfficeNetworkAddressPolicy.CreateRequestAsync(
                    HttpMethod.Get,
                    uri,
                    options.AllowPrivateNetworkAddresses,
                    cancellationToken).ConfigureAwait(false)
                : new HttpRequestMessage(HttpMethod.Get, uri);
            Task<HttpResponseMessage> sendOperation = httpClient.SendAsync(
                request,
                HttpCompletionOption.ResponseHeadersRead,
                cancellationToken);
            try {
                HttpResponseMessage response = await WaitForOperationAsync(
                    sendOperation,
                    cancellationToken,
                    lateResponse => lateResponse.Dispose()).ConfigureAwait(false);
                request.Dispose();
                request = null;
                return response;
            } catch (OperationCanceledException) when (!sendOperation.IsCompleted) {
                DisposeAfterCompletion(sendOperation, request!);
                request = null;
                throw;
            }
        } catch (Exception exception) when (OfficeNetworkAddressPolicy.TryGetRejection(exception, out OfficeNetworkAddressException rejection)) {
            throw ReaderWebUriPolicy.Rejected(uri, rejection);
        } finally {
            request?.Dispose();
        }
    }

    private static bool HasDefaultCredentials(HttpClient httpClient) {
        HttpRequestHeaders headers = httpClient.DefaultRequestHeaders;
        return headers.Authorization != null ||
            headers.Contains("Proxy-Authorization") ||
            headers.Contains("Cookie");
    }

    private static bool IsSameOrigin(Uri left, Uri right) =>
        string.Equals(left.Scheme, right.Scheme, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(left.IdnHost, right.IdnHost, StringComparison.OrdinalIgnoreCase) &&
        left.Port == right.Port;

    private static bool IsRedirect(HttpStatusCode statusCode) =>
        statusCode == HttpStatusCode.Moved ||
        statusCode == HttpStatusCode.Redirect ||
        statusCode == HttpStatusCode.SeeOther ||
        statusCode == HttpStatusCode.TemporaryRedirect ||
        (int)statusCode == 308;

    private static async Task<MemoryStream> ReadBoundedContentAsync(
        HttpContent content,
        long? declaredLength,
        long maxResponseBytes,
        CancellationToken cancellationToken) {
        int capacity = GetInitialBufferCapacity(declaredLength);
        MemoryStream output = ReaderInputLimits.CreateSnapshotStream(capacity);
        try {
            Task<Stream> openOperation = content.ReadAsStreamAsync();
            using Stream input = await WaitForOperationAsync(
                openOperation,
                cancellationToken,
                stream => stream.Dispose()).ConfigureAwait(false);
            byte[] buffer = new byte[64 * 1024];
            long total = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                Task<int> readOperation = input.ReadAsync(buffer, 0, buffer.Length, cancellationToken);
                int read = await WaitForOperationAsync(readOperation, cancellationToken).ConfigureAwait(false);
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
            output.Position = 0;
            return output;
        } catch {
            output.Dispose();
            throw;
        }
    }

    private static async Task<T> WaitForOperationAsync<T>(
        Task<T> operation,
        CancellationToken cancellationToken,
        Action<T>? disposeAbandonedResult = null) {
        if (operation.IsCompleted) {
            return await operation.ConfigureAwait(false);
        }

        var cancellationSignal = new TaskCompletionSource<bool>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        using (cancellationToken.Register(
            state => ((TaskCompletionSource<bool>)state!).TrySetResult(true),
            cancellationSignal)) {
            Task completed = await Task.WhenAny(operation, cancellationSignal.Task).ConfigureAwait(false);
            if (completed == operation) {
                return await operation.ConfigureAwait(false);
            }
        }

        ObserveAbandonedOperation(operation, disposeAbandonedResult);
        throw new OperationCanceledException(cancellationToken);
    }

    private static void ObserveAbandonedOperation<T>(Task<T> operation, Action<T>? disposeResult) {
        _ = operation.ContinueWith(
            completed => {
                if (completed.IsFaulted) {
                    _ = completed.Exception;
                } else if (completed.Status == TaskStatus.RanToCompletion && disposeResult != null) {
                    try {
                        disposeResult(completed.Result);
                    } catch {
                        // Cleanup must not surface on the continuation scheduler.
                    }
                }
            },
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
    }

    private static void DisposeAfterCompletion(Task operation, IDisposable resource) {
        _ = operation.ContinueWith(
            (completed, state) => {
                try {
                    ((IDisposable)state!).Dispose();
                } catch {
                    // Cleanup must not surface on the continuation scheduler.
                }
            },
            resource,
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
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

internal sealed class ReaderWebDownload : IDisposable {
    internal ReaderWebDownload(
        MemoryStream content,
        string sourceName,
        Uri requestUri,
        Uri responseUri,
        int statusCode,
        string? contentType,
        DateTime? lastModifiedUtc) {
        Content = content;
        LengthBytes = content.Length;
        SourceName = sourceName;
        RequestUri = requestUri;
        ResponseUri = responseUri;
        StatusCode = statusCode;
        ContentType = contentType;
        LastModifiedUtc = lastModifiedUtc;
    }

    internal MemoryStream Content { get; }
    internal long LengthBytes { get; }
    internal string SourceName { get; }
    internal Uri RequestUri { get; }
    internal Uri ResponseUri { get; }
    internal int StatusCode { get; }
    internal string? ContentType { get; }
    internal DateTime? LastModifiedUtc { get; }

    internal void ApplyTransportMetadata(
        OfficeDocumentReadResult result,
        ReaderWebOptions options,
        bool computeHashes) {
        string sourceKey = options.SourceKeySelector == null
            ? FormatUri(ResponseUri, includeQuery: true)
            : options.SourceKeySelector(ResponseUri);
        if (string.IsNullOrWhiteSpace(sourceKey)) {
            throw new InvalidOperationException("The web source key selector returned an empty identity key.");
        }
        string sourceId = DocumentReaderEngine.BuildPortableSourceId("web:" + sourceKey);
        foreach (ReaderChunk chunk in result.Chunks) {
            chunk.Id = sourceId + ":" + chunk.Id;
        }
        DocumentReaderEngine.ApplyExternalSourceMetadata(
            result,
            sourceId,
            LastModifiedUtc,
            LengthBytes,
            computeHashes);
        result.CapabilitiesUsed = result.CapabilitiesUsed
            .Concat(new[] { ReaderWebTransport.CapabilityId })
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        var metadata = new List<OfficeDocumentMetadataEntry> {
            Metadata("reader-web-request-uri", "RequestUri", FormatUri(RequestUri, options.IncludeQueryInMetadata), "uri"),
            Metadata("reader-web-response-uri", "ResponseUri", FormatUri(ResponseUri, options.IncludeQueryInMetadata), "uri"),
            Metadata("reader-web-status-code", "StatusCode", StatusCode.ToString(CultureInfo.InvariantCulture), "number"),
            Metadata("reader-web-length", "LengthBytes", LengthBytes.ToString(CultureInfo.InvariantCulture), "number")
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

    public void Dispose() {
        Content.Dispose();
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
