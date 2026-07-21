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
        _ownsClient = session.Options.HttpClient == null;
        _client = session.Options.HttpClient ?? new HttpClient();
        if (_ownsClient) _client.Timeout = Timeout.InfiniteTimeSpan;
    }

    internal async Task<T> SendJsonAsync<T>(HttpMethod method, string relativeUri, object? payload, ConfluenceRequestSafety safety, CancellationToken cancellationToken) {
        ConfluenceHttpResponse response = await SendAsync(
            method,
            relativeUri,
            payload == null ? null : (() => new StringContent(JsonSerializer.Serialize(payload, JsonOptions), Encoding.UTF8, "application/json")),
            null,
            safety,
            cancellationToken).ConfigureAwait(false);
        if (response.Body.Length == 0) return default!;
        T? value = JsonSerializer.Deserialize<T>(response.Body, JsonOptions);
        return value == null ? throw new InvalidOperationException("Confluence returned an empty or invalid JSON response.") : value;
    }

    internal Task<ConfluenceHttpResponse> SendRawAsync(HttpMethod method, string relativeUri, CancellationToken cancellationToken) =>
        SendAsync(method, relativeUri, null, null, ConfluenceRequestSafety.SafeToRetry, cancellationToken);

    internal Task<ConfluenceHttpResponse> SendMultipartAsync(string relativeUri, Func<HttpContent> contentFactory, CancellationToken cancellationToken) =>
        SendAsync(HttpMethod.Put, relativeUri, contentFactory, request => request.Headers.TryAddWithoutValidation("X-Atlassian-Token", "nocheck"), ConfluenceRequestSafety.NonIdempotent, cancellationToken);

    public void Dispose() {
        if (_disposed) return;
        if (_ownsClient) _client.Dispose();
        _disposed = true;
    }

    private async Task<ConfluenceHttpResponse> SendAsync(
        HttpMethod method,
        string relativeUri,
        Func<HttpContent>? contentFactory,
        Action<HttpRequestMessage>? configure,
        ConfluenceRequestSafety safety,
        CancellationToken cancellationToken) {
        ThrowIfDisposed();
        Uri uri = ResolveUri(relativeUri);
        int attempt = 0;
        while (true) {
            using var request = new HttpRequestMessage(method, uri);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.UserAgent.ParseAdd(BuildUserAgent(_session.Options.ApplicationName));
            request.Content = contentFactory?.Invoke();
            configure?.Invoke(request);
            await _session.CredentialSource.ApplyAsync(request, cancellationToken).ConfigureAwait(false);

            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(_session.Options.RequestTimeout);
            HttpResponseMessage response = await _client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, timeout.Token).ConfigureAwait(false);
            if (ShouldRetry(response.StatusCode, safety, attempt)) {
                TimeSpan delay = GetRetryDelay(response, attempt);
                response.Dispose();
                attempt++;
                await Task.Delay(delay, cancellationToken).ConfigureAwait(false);
                continue;
            }

            using (response) {
                byte[] body = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                if (!response.IsSuccessStatusCode) {
                    throw new ConfluenceApiException(response.StatusCode, method.Method, uri, Encoding.UTF8.GetString(body));
                }
                var headers = response.Headers.Concat(response.Content.Headers)
                    .GroupBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => (IReadOnlyList<string>)group.SelectMany(item => item.Value).ToArray(), StringComparer.OrdinalIgnoreCase);
                return new ConfluenceHttpResponse(response.StatusCode, body, headers);
            }
        }
    }

    private bool ShouldRetry(HttpStatusCode statusCode, ConfluenceRequestSafety safety, int attempt) {
        if (safety != ConfluenceRequestSafety.SafeToRetry || attempt >= _session.Options.MaxRetryCount) return false;
        int code = (int)statusCode;
        return code == 429 || code == 500 || code == 502 || code == 503 || code == 504;
    }

    private TimeSpan GetRetryDelay(HttpResponseMessage response, int attempt) {
        if (response.Headers.RetryAfter?.Delta is TimeSpan delta && delta > TimeSpan.Zero) return Clamp(delta);
        if (response.Headers.RetryAfter?.Date is DateTimeOffset date) {
            TimeSpan until = date - DateTimeOffset.UtcNow;
            if (until > TimeSpan.Zero) return Clamp(until);
        }
        double multiplier = Math.Pow(2, attempt);
        return Clamp(TimeSpan.FromMilliseconds(_session.Options.RetryBaseDelay.TotalMilliseconds * multiplier));
    }

    private TimeSpan Clamp(TimeSpan delay) => delay > _session.Options.RetryMaxDelay ? _session.Options.RetryMaxDelay : delay;

    private Uri ResolveUri(string relativeUri) {
        if (string.IsNullOrWhiteSpace(relativeUri)) throw new ArgumentException("Confluence request URI is required.", nameof(relativeUri));
        if (Uri.TryCreate(relativeUri, UriKind.Absolute, out Uri? absolute)) {
            if (!string.Equals(absolute.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Confluence requests must use HTTPS.");
            return absolute;
        }
        return new Uri(_session.Options.SiteUri, relativeUri.StartsWith("/", StringComparison.Ordinal) ? relativeUri : "/" + relativeUri);
    }

    private static string BuildUserAgent(string? applicationName) {
        string value = string.IsNullOrWhiteSpace(applicationName) ? "OfficeIMO" : applicationName!.Trim();
        return value.Replace(' ', '-') + "/3.0";
    }

    private void ThrowIfDisposed() {
        if (_disposed) throw new ObjectDisposedException(nameof(ConfluenceHttpTransport));
    }
}
