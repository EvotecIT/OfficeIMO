using System.Net;

namespace OfficeIMO.Html.Runtime.Worker;

// Document resources and fetch share transport, authority, concurrency and cumulative budgets.
internal sealed class RuntimeResourceLoader : IDisposable {
    private readonly HtmlRuntimeResourcePolicy _policy;
    private readonly Dictionary<string, HtmlRuntimeResource> _supplied;
    private readonly RuntimeResourceBudget _budget;
    private Dictionary<string, HtmlRuntimeResource> _loaded => _budget.Loaded;
    private readonly HashSet<string> _origins;
    private readonly string _documentOrigin;
    private readonly SemaphoreSlim _concurrency;
    private readonly CancellationTokenSource _lifetime = new();
    private readonly HttpClient _client;
    private object _sync => _budget.Sync;

    internal RuntimeResourceLoader(HtmlScriptRequest options, RuntimeResourceBudget? budget = null) {
        _budget = budget ?? new RuntimeResourceBudget(options);
        _policy = options.ResourcePolicy;
        _supplied = options.Resources.ToDictionary(resource => HtmlRuntimeResourcePolicy.Key(resource.Url), StringComparer.Ordinal);
        _documentOrigin = HtmlRuntimeResourcePolicy.Origin(options.DocumentUrl);
        _origins = _budget.Origins;
        _concurrency = _budget.Concurrency;
        _client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false, UseCookies = false, UseProxy = false,
            AutomaticDecompression = DecompressionMethods.None, MaxResponseHeadersLength = 32 }) { Timeout = System.Threading.Timeout.InfiniteTimeSpan };
    }

    internal IReadOnlyList<HtmlRuntimeResource> Capture() { lock (_sync) return _loaded.Values.ToArray(); }
    internal Task<HtmlRuntimeResource> LoadAsync(Uri url, CancellationToken token) => LoadAsync(url, null, null, token);
    internal Task<HtmlRuntimeResource> FetchAsync(Uri url, RuntimeFetchRequest request, CancellationToken token) => LoadAsync(url, request, request.Validate(_policy), token);

    private async Task<HtmlRuntimeResource> LoadAsync(Uri requestedUrl, RuntimeFetchRequest? fetch, byte[]? body, CancellationToken token) {
        _budget.BeginOperation();
        using var deadline = new CancellationTokenSource(_policy.Timeout);
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(token, deadline.Token, _lifetime.Token);
        bool admitted = false;
        try {
            await _concurrency.WaitAsync(operation.Token).ConfigureAwait(false);
            admitted = true;
            string responseFragment = requestedUrl.Fragment;
            requestedUrl = new Uri(HtmlRuntimeResourcePolicy.Key(requestedUrl));
            var currentUrl = requestedUrl;
            string method = fetch?.Method ?? "GET";
            var headers = new Dictionary<string, string>(fetch?.Headers ?? new(), StringComparer.OrdinalIgnoreCase);
            int redirects = 0;
            bool cors = false;
            while (true) {
                operation.Token.ThrowIfCancellationRequested();
                CheckOrigin(currentUrl);
                bool crossOrigin = HtmlRuntimeResourcePolicy.Origin(currentUrl) != _documentOrigin;
                if (fetch != null) {
                    if (crossOrigin && fetch.Mode == "same-origin") throw new HtmlScriptRuntimeException("Cross-origin fetch is forbidden in same-origin mode.");
                    cors |= crossOrigin;
                    if (cors) await PreflightAsync(currentUrl, method, headers, operation.Token).ConfigureAwait(false);
                }
                var outgoing = new Dictionary<string, string>(headers, StringComparer.OrdinalIgnoreCase);
                if (fetch != null && (cors || method is not ("GET" or "HEAD"))) outgoing["Origin"] = _documentOrigin;
                HtmlRuntimeResource response = await SendAsync(currentUrl, method, outgoing, body, operation.Token).ConfigureAwait(false);
                if (fetch != null && cors) RuntimeFetchCors.Check(response.Headers, _documentOrigin);
                if (response.RedirectCount != 0 || response.FinalUrl != currentUrl) {
                    // A supplied final response has no redirect response headers to check for fetch.
                    if (fetch != null) throw new HtmlScriptRuntimeException("Fetch requires explicit redirect responses in supplied resources.");
                    if ((long)redirects + response.RedirectCount > _policy.MaxRedirects) throw new HtmlScriptRuntimeException("Resource redirect budget exceeded.");
                    redirects += response.RedirectCount;
                }
                if (response.StatusCode is 301 or 302 or 303 or 307 or 308) {
                    if (fetch?.Redirect == "error") throw new HtmlScriptRuntimeException("Fetch redirect mode forbids redirects.");
                }
                if (response.StatusCode is 301 or 302 or 303 or 307 or 308 && response.Headers.TryGetValue("Location", out string? location)) {
                    if (++redirects > _policy.MaxRedirects) throw new HtmlScriptRuntimeException("Resource redirect budget exceeded.");
                    Uri next = new Uri(currentUrl, location);
                    // Redirects inherit the current fragment unless Location supplies
                    // one explicitly (including an empty '#'). Never send it over HTTP.
                    if (location.Contains('#')) responseFragment = next.Fragment;
                    CheckOrigin(next);
                    if (fetch != null && HtmlRuntimeResourcePolicy.Origin(next) != HtmlRuntimeResourcePolicy.Origin(currentUrl)) {
                        // A second origin change requires redirect-tainted origin semantics.
                        // Refuse this unqualified case before contacting another authority.
                        if (cors) throw new HtmlScriptRuntimeException("Fetch redirects between distinct cross-origin authorities are not supported.");
                        headers.Remove("Authorization");
                    }
                    if ((response.StatusCode is 301 or 302 && method == "POST") || (response.StatusCode == 303 && method is not ("GET" or "HEAD"))) {
                        method = "GET"; body = null;
                        foreach (string name in new[] { "Content-Encoding", "Content-Language", "Content-Location", "Content-Type" }) headers.Remove(name);
                    }
                    currentUrl = new Uri(HtmlRuntimeResourcePolicy.Key(next));
                    continue;
                }
                var finalUrl = new Uri(HtmlRuntimeResourcePolicy.Key(response.FinalUrl) + responseFragment);
                var result = new HtmlRuntimeResource(requestedUrl, response.Buffer, response.ContentType, response.StatusCode, finalUrl, redirects, response.Headers, response.StatusText);
                // A POST response at an image/script URL must not replace its retained GET asset.
                if (fetch == null || fetch.Method == "GET") { lock (_sync) _loaded[HtmlRuntimeResourcePolicy.Key(requestedUrl)] = result; }
                return result;
            }
        } catch (OperationCanceledException) when (deadline.IsCancellationRequested && !token.IsCancellationRequested && !_lifetime.IsCancellationRequested) {
            throw new HtmlScriptRuntimeException("The resource load exceeded its deadline.");
        } finally {
            if (admitted) _concurrency.Release();
            _budget.EndOperation();
        }
    }

    private async Task PreflightAsync(Uri url, string method, Dictionary<string, string> headers, CancellationToken token) {
        string[] unsafeHeaders = RuntimeFetchCors.UnsafeHeaders(headers);
        if (method is "GET" or "HEAD" or "POST" && unsafeHeaders.Length == 0) return;
        var preflightHeaders = new Dictionary<string, string> { ["Origin"] = _documentOrigin, ["Access-Control-Request-Method"] = method };
        if (unsafeHeaders.Length != 0) preflightHeaders["Access-Control-Request-Headers"] = string.Join(",", unsafeHeaders);
        var response = await SendAsync(url, "OPTIONS", preflightHeaders, null, token).ConfigureAwait(false);
        RuntimeFetchCors.CheckPreflight(response, _documentOrigin, method, unsafeHeaders);
    }

    private async Task<HtmlRuntimeResource> SendAsync(Uri url, string method, IReadOnlyDictionary<string, string> headers, byte[]? body, CancellationToken token) {
        CheckOrigin(url);
        if (Interlocked.Increment(ref _budget.Requests) > _policy.MaxRequests) throw new HtmlScriptRuntimeException("Resource request budget exceeded.");
        if (_supplied.TryGetValue(HtmlRuntimeResourcePolicy.Key(url), out var supplied) && method is "GET" or "HEAD") {
            CheckOrigin(supplied.FinalUrl);
            if (supplied.RedirectCount > _policy.MaxRedirects) throw new HtmlScriptRuntimeException("Resource redirect budget exceeded.");
            if (Interlocked.Add(ref _budget.Requests, supplied.RedirectCount) > _policy.MaxRequests) throw new HtmlScriptRuntimeException("Resource request budget exceeded.");
            ReserveBytes(method == "HEAD" ? 0 : supplied.Length);
            var suppliedHeaders = new Dictionary<string, string>(supplied.Headers, StringComparer.OrdinalIgnoreCase) { ["Content-Type"] = supplied.ContentType };
            return new HtmlRuntimeResource(url, method == "HEAD" ? Array.Empty<byte>() : supplied.Buffer, supplied.ContentType, supplied.StatusCode, new Uri(HtmlRuntimeResourcePolicy.Key(supplied.FinalUrl)), supplied.RedirectCount, suppliedHeaders, supplied.StatusText);
        }
        if (!_policy.AllowNetwork) throw new HtmlScriptRuntimeException("The resource was not supplied and network loading is disabled.");
        if (body != null) {
            lock (_sync) {
                if (body.LongLength > _policy.MaxTotalRequestBytes - _budget.SentBytes) throw new HtmlScriptRuntimeException("Total fetch request body byte budget exceeded.");
                _budget.SentBytes += body.LongLength;
            }
        }
        using var request = new HttpRequestMessage(new HttpMethod(method), url);
        if (body != null) request.Content = new ByteArrayContent(body);
        foreach (var header in headers) {
            if (!request.Headers.TryAddWithoutValidation(header.Key, header.Value)) {
                request.Content ??= new ByteArrayContent(Array.Empty<byte>());
                request.Content.Headers.TryAddWithoutValidation(header.Key, header.Value);
            }
        }
        request.Headers.TryAddWithoutValidation("Accept-Encoding", "identity");
        using var response = await _client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, token).ConfigureAwait(false);
        var responseHeaders = response.Headers.Concat(response.Content.Headers).ToDictionary(h => h.Key, h => string.Join(", ", h.Value), StringComparer.OrdinalIgnoreCase);
        if (response.Content.Headers.ContentEncoding.Any(value => !value.Equals("identity", StringComparison.OrdinalIgnoreCase)))
            throw new HtmlScriptRuntimeException("Compressed runtime responses are not supported by this transport profile.");
        if (method != "HEAD" && response.Content.Headers.ContentLength > _policy.MaxResourceBytes) throw new HtmlScriptRuntimeException("Resource response byte budget exceeded.");
        await using var input = await response.Content.ReadAsStreamAsync(token).ConfigureAwait(false);
        using var output = new MemoryStream();
        var buffer = new byte[16 * 1024];
        int read;
        while (method != "HEAD" && (read = await input.ReadAsync(buffer, token).ConfigureAwait(false)) != 0) {
            if (output.Length + read > _policy.MaxResourceBytes) throw new HtmlScriptRuntimeException("Resource response byte budget exceeded.");
            ReserveBytes(read);
            output.Write(buffer, 0, read);
        }
        token.ThrowIfCancellationRequested();
        return new HtmlRuntimeResource(url, output.ToArray(), response.Content.Headers.ContentType?.ToString() ?? "application/octet-stream", (int)response.StatusCode, headers: responseHeaders, statusText: response.ReasonPhrase ?? "");
    }

    private void CheckOrigin(Uri url) {
        if (!_origins.Contains(HtmlRuntimeResourcePolicy.Origin(url))) throw new HtmlScriptRuntimeException("The resource origin is not allowed.");
    }
    private void ReserveBytes(long count) {
        lock (_sync) {
            if (count > _policy.MaxTotalBytes - _budget.ReceivedBytes) throw new HtmlScriptRuntimeException("Total resource response byte budget exceeded.");
            _budget.ReceivedBytes += count;
        }
    }
    public void Dispose() {
        _lifetime.Cancel();
        _client.Dispose();
        // Outstanding provider loads may still observe the token and release their slot.
    }
}
