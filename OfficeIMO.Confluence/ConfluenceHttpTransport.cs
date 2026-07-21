using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Confluence;

internal enum ConfluenceRequestSafety {
    SafeToRetry,
    NonIdempotent,
}

/// <summary>An error returned by Confluence Cloud.</summary>
public sealed class ConfluenceApiException : Exception {
    internal ConfluenceApiException(HttpStatusCode statusCode, string method, Uri requestUri, string responseBody)
        : base(BuildMessage(statusCode, method, requestUri, responseBody)) {
        StatusCode = statusCode;
        RequestMethod = method;
        RequestUri = requestUri;
        ResponseBody = responseBody;
    }

    public HttpStatusCode StatusCode { get; }
    public string RequestMethod { get; }
    public Uri RequestUri { get; }
    public string ResponseBody { get; }

    private static string BuildMessage(HttpStatusCode statusCode, string method, Uri requestUri, string responseBody) {
        string compact = string.IsNullOrWhiteSpace(responseBody) ? "No response body." : responseBody.Trim();
        if (compact.Length > 2048) compact = compact.Substring(0, 2048) + "...";
        return "Confluence " + method + " " + requestUri.PathAndQuery + " failed with HTTP " + (int)statusCode + ": " + compact;
    }
}

internal sealed class ConfluenceHttpResponse {
    internal ConfluenceHttpResponse(HttpStatusCode statusCode, byte[] body, IReadOnlyDictionary<string, IReadOnlyList<string>> headers) {
        StatusCode = statusCode;
        Body = body;
        Headers = headers;
    }
    internal HttpStatusCode StatusCode { get; }
    internal byte[] Body { get; }
    internal IReadOnlyDictionary<string, IReadOnlyList<string>> Headers { get; }
}

internal sealed class ConfluenceJsonResponse<T> {
    internal ConfluenceJsonResponse(T value, IReadOnlyDictionary<string, IReadOnlyList<string>> headers) {
        Value = value;
        Headers = headers;
    }
    internal T Value { get; }
    internal IReadOnlyDictionary<string, IReadOnlyList<string>> Headers { get; }
}

internal sealed class ConfluenceHttpTransport : IDisposable {
    internal static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions {
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
    };

    private readonly ConfluenceSession _session;
    private readonly HttpClient _client;
    private readonly bool _ownsClient;
    private bool _disposed;

    internal ConfluenceHttpTransport(ConfluenceSession session) {
        _session = session ?? throw new ArgumentNullException(nameof(session));
        _ownsClient = session.RuntimeOptions.HttpClient == null;
        _client = session.RuntimeOptions.HttpClient ?? new HttpClient();
        if (_ownsClient) _client.Timeout = Timeout.InfiniteTimeSpan;
    }

    internal async Task<T> SendJsonAsync<T>(HttpMethod method, string relativeUri, object? payload, ConfluenceRequestSafety safety, CancellationToken cancellationToken) {
        ConfluenceJsonResponse<T> response = await SendJsonWithHeadersAsync<T>(method, relativeUri, payload, safety, cancellationToken).ConfigureAwait(false);
        return response.Value;
    }

    internal async Task<ConfluenceJsonResponse<T>> SendJsonWithHeadersAsync<T>(HttpMethod method, string relativeUri, object? payload, ConfluenceRequestSafety safety, CancellationToken cancellationToken) {
        ConfluenceHttpResponse response = await SendAsync(
            method,
            relativeUri,
            payload == null ? null : (() => new StringContent(JsonSerializer.Serialize(payload, JsonOptions), Encoding.UTF8, "application/json")),
            null,
            safety,
            cancellationToken).ConfigureAwait(false);
        if (response.Body.Length == 0) return new ConfluenceJsonResponse<T>(default!, response.Headers);
        T? value = JsonSerializer.Deserialize<T>(response.Body, JsonOptions);
        return value == null
            ? throw new InvalidOperationException("Confluence returned an empty or invalid JSON response.")
            : new ConfluenceJsonResponse<T>(value, response.Headers);
    }

    internal Task<ConfluenceHttpResponse> SendRawAsync(HttpMethod method, string relativeUri, CancellationToken cancellationToken) =>
        SendAsync(method, relativeUri, null, null, ConfluenceRequestSafety.SafeToRetry, cancellationToken);

    internal Task<ConfluenceHttpResponse> SendRawAsync(HttpMethod method, string relativeUri, ConfluenceRequestSafety safety, CancellationToken cancellationToken) =>
        SendAsync(method, relativeUri, null, null, safety, cancellationToken);

    internal Task SendToStreamAsync(HttpMethod method, string relativeUri, Stream destination, CancellationToken cancellationToken) {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("Attachment destination stream must be writable.", nameof(destination));
        return SendToStreamCoreAsync(method, relativeUri, destination, cancellationToken);
    }

    internal Task<ConfluenceHttpResponse> SendMultipartAsync(string relativeUri, Func<HttpContent> contentFactory, CancellationToken cancellationToken) =>
        SendAsync(HttpMethod.Put, relativeUri, contentFactory, request => request.Headers.TryAddWithoutValidation("X-Atlassian-Token", "nocheck"), ConfluenceRequestSafety.NonIdempotent, cancellationToken);

    public void Dispose() {
        if (_disposed) return;
        if (_ownsClient) _client.Dispose();
        _disposed = true;
    }

    private async Task SendToStreamCoreAsync(HttpMethod method, string relativeUri, Stream destination, CancellationToken cancellationToken) {
        await SendAsync(
                method,
                relativeUri,
                null,
                null,
                ConfluenceRequestSafety.SafeToRetry,
                cancellationToken,
                async (response, token) => {
                    using Stream source = await ReadAsStreamAsync(response.Content, token).ConfigureAwait(false);
                    await source.CopyToAsync(destination, 81920, token).ConfigureAwait(false);
                    return true;
                })
            .ConfigureAwait(false);
    }

    private async Task<ConfluenceHttpResponse> SendAsync(
        HttpMethod method,
        string relativeUri,
        Func<HttpContent>? contentFactory,
        Action<HttpRequestMessage>? configure,
        ConfluenceRequestSafety safety,
        CancellationToken cancellationToken) => await SendAsync(
            method,
            relativeUri,
            contentFactory,
            configure,
            safety,
            cancellationToken,
            async (response, token) => new ConfluenceHttpResponse(
                response.StatusCode,
                await ReadBodyAsync(response.Content, token).ConfigureAwait(false),
                CaptureHeaders(response)))
        .ConfigureAwait(false);

    private async Task<T> SendAsync<T>(
        HttpMethod method,
        string relativeUri,
        Func<HttpContent>? contentFactory,
        Action<HttpRequestMessage>? configure,
        ConfluenceRequestSafety safety,
        CancellationToken cancellationToken,
        Func<HttpResponseMessage, CancellationToken, Task<T>> readSuccess) {
        ThrowIfDisposed();
        Uri uri = ResolveUri(relativeUri);
        int attempt = 0;
        while (true) {
            using var request = new HttpRequestMessage(method, uri);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.UserAgent.ParseAdd(BuildUserAgent(_session.RuntimeOptions.ApplicationName));
            request.Content = contentFactory?.Invoke();
            configure?.Invoke(request);
            await _session.CredentialSource.ApplyAsync(request, cancellationToken).ConfigureAwait(false);

            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(_session.RuntimeOptions.RequestTimeout);
            try {
                using HttpResponseMessage response = await _client
                    .SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token)
                    .ConfigureAwait(false);
                if (ShouldRetry(response.StatusCode, safety, attempt)) {
                    TimeSpan delay = GetRetryDelay(response, attempt);
                    attempt++;
                    await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
                    continue;
                }
                if (!response.IsSuccessStatusCode) {
                    byte[] body = await ReadBodyAsync(response.Content, timeout.Token).ConfigureAwait(false);
                    throw new ConfluenceApiException(response.StatusCode, method.Method, uri, Encoding.UTF8.GetString(body));
                }
                return await readSuccess(response, timeout.Token).ConfigureAwait(false);
            } catch (HttpRequestException) when (CanRetry(safety, attempt)) {
                TimeSpan delay = GetRetryDelay(attempt);
                attempt++;
                await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
                continue;
            } catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested && CanRetry(safety, attempt)) {
                TimeSpan delay = GetRetryDelay(attempt);
                attempt++;
                await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
                continue;
            }
        }
    }

    private bool ShouldRetry(HttpStatusCode statusCode, ConfluenceRequestSafety safety, int attempt) {
        if (!CanRetry(safety, attempt)) return false;
        int code = (int)statusCode;
        return code == 429 || code == 500 || code == 502 || code == 503 || code == 504;
    }

    private bool CanRetry(ConfluenceRequestSafety safety, int attempt) =>
        safety == ConfluenceRequestSafety.SafeToRetry && attempt < _session.RuntimeOptions.MaxRetryCount;

    private TimeSpan GetRetryDelay(HttpResponseMessage response, int attempt) {
        if (response.Headers.RetryAfter?.Delta is TimeSpan delta && delta > TimeSpan.Zero) return Clamp(delta);
        if (response.Headers.RetryAfter?.Date is DateTimeOffset date) {
            TimeSpan until = date - DateTimeOffset.UtcNow;
            if (until > TimeSpan.Zero) return Clamp(until);
        }
        double multiplier = Math.Pow(2, attempt);
        return Clamp(TimeSpan.FromMilliseconds(_session.RuntimeOptions.RetryBaseDelay.TotalMilliseconds * multiplier));
    }

    private TimeSpan GetRetryDelay(int attempt) {
        double multiplier = Math.Pow(2, attempt);
        return Clamp(TimeSpan.FromMilliseconds(_session.RuntimeOptions.RetryBaseDelay.TotalMilliseconds * multiplier));
    }

    private TimeSpan Clamp(TimeSpan delay) => delay > _session.RuntimeOptions.RetryMaxDelay ? _session.RuntimeOptions.RetryMaxDelay : delay;

    private Uri ResolveUri(string relativeUri) {
        if (string.IsNullOrWhiteSpace(relativeUri)) throw new ArgumentException("Confluence request URI is required.", nameof(relativeUri));
        Uri resolved = Uri.TryCreate(relativeUri, UriKind.Absolute, out Uri? absolute)
            ? absolute
            : new Uri(_session.ApiBaseUri, relativeUri.TrimStart('/'));
        ValidateResolvedUri(resolved);
        return resolved;
    }

    private void ValidateResolvedUri(Uri uri) {
        if (!string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Confluence requests must use HTTPS.");
        if (!string.IsNullOrEmpty(uri.UserInfo)) throw new InvalidOperationException("Confluence request URIs cannot contain embedded credentials.");
        Uri expected = _session.ApiBaseUri;
        if (!string.Equals(uri.Scheme, expected.Scheme, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(uri.Host, expected.Host, StringComparison.OrdinalIgnoreCase) ||
            uri.Port != expected.Port) {
            throw new InvalidOperationException("Confluence requests cannot leave the configured API origin.");
        }
        string expectedPath = expected.AbsolutePath.EndsWith("/", StringComparison.Ordinal) ? expected.AbsolutePath : expected.AbsolutePath + "/";
        if (!uri.AbsolutePath.StartsWith(expectedPath, StringComparison.Ordinal)) {
            throw new InvalidOperationException("Confluence requests cannot leave the configured API path.");
        }
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<string>> CaptureHeaders(HttpResponseMessage response) =>
        response.Headers.Concat(response.Content.Headers)
            .GroupBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => (IReadOnlyList<string>)group.SelectMany(item => item.Value).ToArray(), StringComparer.OrdinalIgnoreCase);

    private static async Task<byte[]> ReadBodyAsync(HttpContent content, CancellationToken cancellationToken) {
        using Stream source = await ReadAsStreamAsync(content, cancellationToken).ConfigureAwait(false);
        using var buffer = new MemoryStream();
        await source.CopyToAsync(buffer, 81920, cancellationToken).ConfigureAwait(false);
        return buffer.ToArray();
    }

    private static Task<Stream> ReadAsStreamAsync(HttpContent content, CancellationToken cancellationToken) {
#if NET8_0_OR_GREATER
        return content.ReadAsStreamAsync(cancellationToken);
#else
        cancellationToken.ThrowIfCancellationRequested();
        return content.ReadAsStreamAsync();
#endif
    }

    private static string BuildUserAgent(string? applicationName) {
        string value = string.IsNullOrWhiteSpace(applicationName) ? "OfficeIMO" : applicationName!.Trim();
        return value.Replace(' ', '-') + "/3.0";
    }

    private void ThrowIfDisposed() {
        if (_disposed) throw new ObjectDisposedException(nameof(ConfluenceHttpTransport));
    }
}
