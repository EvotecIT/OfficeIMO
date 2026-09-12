using System.Net;

namespace OfficeIMO.Html.Runtime.Worker;

// One transport/policy owner for native document loads and subsequent runtime fetch bindings.
internal sealed class RuntimeResourceLoader : IDisposable {
    private readonly HtmlRuntimeResourcePolicy _policy;
    private readonly Dictionary<string, HtmlRuntimeResource> _supplied;
    private readonly Dictionary<string, HtmlRuntimeResource> _loaded = new(StringComparer.Ordinal);
    private readonly HashSet<string> _origins;
    private readonly SemaphoreSlim _concurrency;
    private readonly CancellationTokenSource _lifetime = new();
    private readonly HttpClient _client;
    private readonly object _sync = new();
    private long _bytes;
    private long _requests;

    internal RuntimeResourceLoader(HtmlScriptRequest options) {
        _policy = options.ResourcePolicy;
        _supplied = options.Resources.ToDictionary(resource => HtmlRuntimeResourcePolicy.Key(resource.Url), StringComparer.Ordinal);
        _origins = new HashSet<string>(_policy.AllowedOrigins.Select(HtmlRuntimeResourcePolicy.Origin), StringComparer.OrdinalIgnoreCase) {
            HtmlRuntimeResourcePolicy.Origin(options.DocumentUrl)
        };
        _concurrency = new SemaphoreSlim(_policy.MaxConcurrentRequests);
        _client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false, UseCookies = false, UseProxy = false,
            AutomaticDecompression = DecompressionMethods.None, MaxResponseHeadersLength = 32 }) { Timeout = System.Threading.Timeout.InfiniteTimeSpan };
    }

    internal IReadOnlyList<HtmlRuntimeResource> Capture() {
        lock (_sync) return _loaded.Values.ToArray();
    }

    internal async Task<HtmlRuntimeResource> LoadAsync(Uri requestedUrl, CancellationToken token) {
        using var deadline = new CancellationTokenSource(_policy.Timeout);
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(token, deadline.Token, _lifetime.Token);
        bool admitted = false;
        try {
            await _concurrency.WaitAsync(operation.Token).ConfigureAwait(false);
            admitted = true;
            requestedUrl = new Uri(HtmlRuntimeResourcePolicy.Key(requestedUrl));
            var currentUrl = requestedUrl;
            int redirects = 0;
            while (true) {
                operation.Token.ThrowIfCancellationRequested();
                CheckOrigin(currentUrl);
                ReserveRequest();
                if (_supplied.TryGetValue(HtmlRuntimeResourcePolicy.Key(currentUrl), out var supplied)) {
                    CheckOrigin(supplied.FinalUrl);
                    if ((long)redirects + supplied.RedirectCount > _policy.MaxRedirects) throw new HtmlScriptRuntimeException("Resource redirect budget exceeded.");
                    ReserveRequest(supplied.RedirectCount);
                    ReserveBytes(supplied.Length);
                    return Remember(new HtmlRuntimeResource(requestedUrl, supplied.Buffer, supplied.ContentType, supplied.StatusCode, supplied.FinalUrl, redirects + supplied.RedirectCount));
                }
                if (!_policy.AllowNetwork) throw new HtmlScriptRuntimeException("The resource was not supplied and network loading is disabled.");
                using var request = new HttpRequestMessage(HttpMethod.Get, currentUrl);
                using var response = await _client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, operation.Token).ConfigureAwait(false);
                if (response.StatusCode is HttpStatusCode.MovedPermanently or HttpStatusCode.Redirect or HttpStatusCode.SeeOther or HttpStatusCode.TemporaryRedirect or HttpStatusCode.PermanentRedirect) {
                    if (++redirects > _policy.MaxRedirects) throw new HtmlScriptRuntimeException("Resource redirect budget exceeded.");
                    Uri location = response.Headers.Location ?? throw new HtmlScriptRuntimeException("The resource redirect has no location.");
                    currentUrl = location.IsAbsoluteUri ? location : new Uri(currentUrl, location);
                    continue;
                }
                if (response.Content.Headers.ContentLength > _policy.MaxResourceBytes) throw new HtmlScriptRuntimeException("Resource response byte budget exceeded.");
                await using var input = await response.Content.ReadAsStreamAsync(operation.Token).ConfigureAwait(false);
                using var output = new MemoryStream();
                var buffer = new byte[16 * 1024];
                int read;
                while ((read = await input.ReadAsync(buffer, operation.Token).ConfigureAwait(false)) != 0) {
                    if (output.Length + read > _policy.MaxResourceBytes) throw new HtmlScriptRuntimeException("Resource response byte budget exceeded.");
                    ReserveBytes(read);
                    output.Write(buffer, 0, read);
                }
                operation.Token.ThrowIfCancellationRequested();
                return Remember(new HtmlRuntimeResource(requestedUrl, output.ToArray(), response.Content.Headers.ContentType?.ToString() ?? "application/octet-stream", (int)response.StatusCode, currentUrl, redirects));
            }
        } catch (OperationCanceledException) when (deadline.IsCancellationRequested && !token.IsCancellationRequested && !_lifetime.IsCancellationRequested) {
            throw new HtmlScriptRuntimeException("The resource load exceeded its deadline.");
        } finally { if (admitted) _concurrency.Release(); }
    }

    private void CheckOrigin(Uri url) {
        if (!_origins.Contains(HtmlRuntimeResourcePolicy.Origin(url))) throw new HtmlScriptRuntimeException("The resource origin is not allowed.");
    }

    private void ReserveRequest(int count = 1) {
        if (Interlocked.Add(ref _requests, count) > _policy.MaxRequests) throw new HtmlScriptRuntimeException("Resource request budget exceeded.");
    }

    private void ReserveBytes(long count) {
        lock (_sync) {
            if (count > _policy.MaxTotalBytes - _bytes) throw new HtmlScriptRuntimeException("Total resource response byte budget exceeded.");
            _bytes += count;
        }
    }

    private HtmlRuntimeResource Remember(HtmlRuntimeResource resource) {
        lock (_sync) _loaded[HtmlRuntimeResourcePolicy.Key(resource.Url)] = resource;
        return resource;
    }

    public void Dispose() {
        _lifetime.Cancel();
        _client.Dispose();
        // Outstanding provider loads may still observe this token and release their slot.
        // The worker owns one loader; the process exits immediately after session disposal.
    }
}
