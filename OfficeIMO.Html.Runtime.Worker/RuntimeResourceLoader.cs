using System.Net;

namespace OfficeIMO.Html.Runtime.Worker;

// Document resources and fetch share transport, authority, concurrency and cumulative budgets.
internal sealed class RuntimeResourceLoader : IDisposable {
    internal const string MissingResourceMessage = "The resource was not supplied and network loading is disabled.";
    private readonly HtmlRuntimeResourcePolicy _policy;
    private readonly Dictionary<string, HtmlRuntimeResource> _supplied;
    private readonly Dictionary<string, HtmlRuntimeFetchReplay> _fetchReplays;
    private readonly Dictionary<string, HtmlRuntimeNavigationReplay> _navigationReplays;
    private readonly RuntimeResourceBudget _budget;
    private Dictionary<string, HtmlRuntimeResource> _loaded => _budget.Loaded;
    private readonly HashSet<string> _origins;
    private readonly string _documentOrigin;
    private readonly SemaphoreSlim _concurrency;
    private readonly CancellationTokenSource _lifetime = new();
    private readonly HttpClient _client;
    private readonly RuntimeDiagnostics _diagnostics;
    private readonly bool _discoverNavigationReplays;
    private object _sync => _budget.Sync;

    internal RuntimeResourceLoader(HtmlScriptRequest options, RuntimeResourceBudget? budget, RuntimeDiagnostics diagnostics) {
        _budget = budget ?? new RuntimeResourceBudget(options);
        _policy = options.ResourcePolicy;
        _supplied = options.Resources.ToDictionary(resource => HtmlRuntimeResourcePolicy.Key(resource.Url), StringComparer.Ordinal);
        _fetchReplays = options.FetchReplays.ToDictionary(replay => replay.Identity, StringComparer.Ordinal);
        _navigationReplays = options.NavigationReplays.ToDictionary(replay => replay.Identity, StringComparer.Ordinal);
        _discoverNavigationReplays = options.FailOnNavigationReplayDiscovery || _navigationReplays.Count != 0;
        _documentOrigin = HtmlRuntimeResourcePolicy.Origin(options.DocumentUrl);
        _origins = _budget.Origins;
        _concurrency = _budget.Concurrency;
        _diagnostics = diagnostics;
        _client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false, UseCookies = false, UseProxy = false,
            AutomaticDecompression = DecompressionMethods.None, MaxResponseHeadersLength = 32 }) { Timeout = System.Threading.Timeout.InfiniteTimeSpan };
    }

    internal IReadOnlyList<HtmlRuntimeResource> Capture() { lock (_sync) return _loaded.Values.ToArray(); }
    internal Task<HtmlRuntimeResource> LoadAsync(Uri url, CancellationToken token) => LoadAsync(url, null, null, null, token);
    internal Task<HtmlRuntimeResource> LoadNavigationAsync(HtmlRuntimeNavigationRequest request, CancellationToken token) =>
        _discoverNavigationReplays ? ReplayNavigationAsync(request, token) : LoadAsync(request.Url, token);
    internal Task<HtmlRuntimeResource> FetchAsync(Uri url, RuntimeFetchRequest request, CancellationToken token) {
        byte[]? body = request.Validate(_policy);
        return LoadAsync(url, request, request.ReplayRequest(new Uri(HtmlRuntimeResourcePolicy.Key(url)), body,
            new Uri(_documentOrigin + "/")), body, token);
    }

    private async Task<HtmlRuntimeResource> LoadAsync(Uri requestedUrl, RuntimeFetchRequest? fetch,
        HtmlRuntimeFetchRequest? replayRequest, byte[]? body, CancellationToken token) {
        DateTimeOffset started = DateTimeOffset.UtcNow;
        Uri originalUrl = requestedUrl;
        string originalMethod = fetch?.Method ?? "GET";
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
            HtmlRuntimeFetchReplay? replay = null;
            int replayHop = 0;
            if (replayRequest != null) {
                RuntimeFetchRequest replayFetch = fetch!;
                bool replayCrossOrigin = HtmlRuntimeResourcePolicy.Origin(requestedUrl) != _documentOrigin;
                if (replayCrossOrigin && replayFetch.Mode == "same-origin")
                    throw new HtmlScriptRuntimeException("Cross-origin fetch is forbidden in same-origin mode.");
                HtmlRuntimeFetchDiscovery discovery = NextFetchOccurrence(replayRequest);
                _fetchReplays.TryGetValue(discovery.Identity, out replay);
                bool urlSupply = body == null && headers.Count == 0 && method is "GET" or "HEAD"
                    && _supplied.ContainsKey(HtmlRuntimeResourcePolicy.Key(requestedUrl));
                if (replay == null && !_policy.AllowNetwork && !urlSupply) {
                    ReserveExactRequest(body);
                    _diagnostics.RecordMissingFetch(discovery);
                    bool simpleGet = method == "GET" && body == null && headers.Count == 0;
                    if (simpleGet) _diagnostics.RecordMissingResource(requestedUrl);
                    _diagnostics.Record(HtmlRuntimeEventKind.Policy, "network-access", "blocked", started,
                        url: requestedUrl, method: method, decision: simpleGet ? "network-disabled-replayable-get" : "network-disabled");
                    throw new HtmlScriptRuntimeException(MissingResourceMessage);
                }
            }
            int redirects = 0;
            bool cors = false;
            while (true) {
                operation.Token.ThrowIfCancellationRequested();
                if (replay != null && replayHop >= replay.Hops.Count)
                    throw new HtmlScriptRuntimeException("The dynamic replay ended before the fetch completed.");
                CheckOrigin(currentUrl);
                bool crossOrigin = HtmlRuntimeResourcePolicy.Origin(currentUrl) != _documentOrigin;
                if (fetch != null) {
                    if (crossOrigin && fetch.Mode == "same-origin") throw new HtmlScriptRuntimeException("Cross-origin fetch is forbidden in same-origin mode.");
                    cors |= crossOrigin;
                    if (cors) {
                        if (replay != null) {
                            HtmlRuntimeResource? preflight = replay.Hops[replayHop].PreflightResponse;
                            if (HtmlRuntimeCorsPolicy.RequiresPreflight(method, headers)) {
                                if (preflight == null) throw new HtmlScriptRuntimeException("The dynamic replay omitted its required CORS preflight.");
                                ReserveExactRequest(null);
                                ReserveBytes(preflight.Length);
                                HtmlRuntimeCorsPolicy.CheckPreflight(preflight, _documentOrigin, method, HtmlRuntimeCorsPolicy.UnsafeHeaders(headers));
                            } else if (preflight != null) throw new HtmlScriptRuntimeException("The dynamic replay supplied an unexpected CORS preflight.");
                        } else await PreflightAsync(currentUrl, method, headers, operation.Token).ConfigureAwait(false);
                    } else if (replay?.Hops[replayHop].PreflightResponse != null)
                        throw new HtmlScriptRuntimeException("The dynamic replay supplied a same-origin CORS preflight.");
                }
                var outgoing = new Dictionary<string, string>(headers, StringComparer.OrdinalIgnoreCase);
                if (fetch != null && (cors || method is not ("GET" or "HEAD"))) outgoing["Origin"] = _documentOrigin;
                bool allowUrlSupply = fetch == null || body == null && fetch.Headers.Count == 0 && method is "GET" or "HEAD";
                HtmlRuntimeResource response;
                if (replay != null) {
                    ReserveExactRequest(body);
                    response = replay.Hops[replayHop++].Response;
                    if (HtmlRuntimeResourcePolicy.Key(response.Url) != HtmlRuntimeResourcePolicy.Key(currentUrl))
                        throw new HtmlScriptRuntimeException("The dynamic replay hop URL did not match the fetch redirect.");
                    ReserveBytes(method == "HEAD" ? 0 : response.Length);
                    if (replayHop == 1) {
                        _diagnostics.RecordConsumedFetchReplay(replay.Identity);
                        _diagnostics.Record(HtmlRuntimeEventKind.Policy, "fetch-replay", "consumed", started,
                            url: requestedUrl, method: method, decision: "supplied-dynamic-replay", artifactId: replay.Identity);
                    }
                } else response = await SendAsync(currentUrl, method, outgoing, body, allowUrlSupply, operation.Token).ConfigureAwait(false);
                if (fetch != null && cors) HtmlRuntimeCorsPolicy.Check(response.Headers, _documentOrigin);
                if (response.RedirectCount != 0 || response.FinalUrl != currentUrl) {
                    // A supplied final response has no redirect response headers to check for fetch.
                    if (fetch != null) throw new HtmlScriptRuntimeException("Fetch requires explicit redirect responses in supplied resources.");
                    if ((long)redirects + response.RedirectCount > _policy.MaxRedirects) throw new HtmlScriptRuntimeException("Resource redirect budget exceeded.");
                    redirects += response.RedirectCount;
                    _diagnostics.Record(HtmlRuntimeEventKind.Redirect, "resource-redirect", "followed", started,
                        url: response.FinalUrl, method: method, statusCode: response.StatusCode, redirectCount: redirects);
                }
                if (response.StatusCode is 301 or 302 or 303 or 307 or 308) {
                    if (fetch?.Redirect == "error") throw new HtmlScriptRuntimeException("Fetch redirect mode forbids redirects.");
                }
                if (response.StatusCode is 301 or 302 or 303 or 307 or 308 && response.Headers.TryGetValue("Location", out string? location)) {
                    if (++redirects > _policy.MaxRedirects) throw new HtmlScriptRuntimeException("Resource redirect budget exceeded.");
                    Uri next = new Uri(currentUrl, location);
                    _diagnostics.Record(HtmlRuntimeEventKind.Redirect, "resource-redirect", "followed", started,
                        url: next, method: method, statusCode: response.StatusCode, redirectCount: redirects);
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
                if (replay != null && replayHop != replay.Hops.Count)
                    throw new HtmlScriptRuntimeException("The dynamic replay contains responses after the fetch completed.");
                var finalUrl = new Uri(HtmlRuntimeResourcePolicy.Key(response.FinalUrl) + responseFragment);
                var result = new HtmlRuntimeResource(requestedUrl, response.Buffer, response.ContentType, response.StatusCode, finalUrl, redirects, response.Headers, response.StatusText);
                // A POST response at an image/script URL must not replace its retained GET asset.
                if (fetch == null || fetch.Method == "GET") { lock (_sync) _loaded[HtmlRuntimeResourcePolicy.Key(requestedUrl)] = result; }
                _diagnostics.Record(HtmlRuntimeEventKind.Resource, fetch == null ? "resource-load" : "fetch", "success", started,
                    DateTimeOffset.UtcNow - started, url: originalUrl, method: originalMethod, statusCode: result.StatusCode,
                    byteCount: result.Length, redirectCount: result.RedirectCount, decision: "allowed");
                return result;
            }
        } catch (OperationCanceledException) when (deadline.IsCancellationRequested && !token.IsCancellationRequested && !_lifetime.IsCancellationRequested) {
            _diagnostics.Record(HtmlRuntimeEventKind.Resource, fetch == null ? "resource-load" : "fetch", "timeout", started,
                DateTimeOffset.UtcNow - started, url: originalUrl, method: originalMethod, decision: "blocked");
            throw new HtmlScriptRuntimeException("The resource load exceeded its deadline.");
        } catch (Exception error) {
            _diagnostics.Record(HtmlRuntimeEventKind.Resource, fetch == null ? "resource-load" : "fetch", "failure", started,
                DateTimeOffset.UtcNow - started, detail: _diagnostics.IncludeFailureMessages ? error.Message : null,
                url: originalUrl, method: originalMethod, decision: "blocked");
            throw;
        } finally {
            if (admitted) _concurrency.Release();
            _budget.EndOperation();
        }
    }

    private async Task<HtmlRuntimeResource> ReplayNavigationAsync(HtmlRuntimeNavigationRequest request, CancellationToken token) {
        DateTimeOffset started = DateTimeOffset.UtcNow;
        _budget.BeginOperation();
        using var deadline = new CancellationTokenSource(_policy.Timeout);
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(token, deadline.Token, _lifetime.Token);
        bool admitted = false;
        try {
            await _concurrency.WaitAsync(operation.Token).ConfigureAwait(false);
            admitted = true;
            HtmlRuntimeNavigationDiscovery discovery = NextNavigationOccurrence(request);
            if (!_navigationReplays.TryGetValue(discovery.Identity, out HtmlRuntimeNavigationReplay? replay)) {
                ReserveExactRequest(request.Buffer);
                _diagnostics.RecordMissingNavigation(discovery);
                _diagnostics.Record(HtmlRuntimeEventKind.Policy, "navigation-network-access", "blocked", started,
                    url: request.Url, method: request.Method, decision: "network-disabled-exact-navigation");
                throw new HtmlScriptRuntimeException(MissingResourceMessage);
            }

            Uri originalUrl = request.Url;
            string responseFragment = originalUrl.Fragment;
            Uri currentUrl = new(HtmlRuntimeResourcePolicy.Key(originalUrl));
            string method = request.Method;
            byte[]? body = request.Buffer;
            int redirects = 0;
            for (int index = 0; index < replay.Hops.Count; index++) {
                operation.Token.ThrowIfCancellationRequested();
                CheckOrigin(currentUrl);
                HtmlRuntimeResource response = replay.Hops[index].Response;
                if (HtmlRuntimeResourcePolicy.Key(response.Url) != HtmlRuntimeResourcePolicy.Key(currentUrl))
                    throw new HtmlScriptRuntimeException("The navigation replay hop URL did not match the document redirect.");
                ReserveExactRequest(body);
                ReserveBytes(method == "HEAD" ? 0 : response.Length);
                if (index == 0) {
                    _diagnostics.RecordConsumedNavigationReplay(replay.Identity);
                    _diagnostics.Record(HtmlRuntimeEventKind.Policy, "navigation-replay", "consumed", started,
                        url: originalUrl, method: method, decision: "supplied-navigation-replay", artifactId: replay.Identity);
                }
                bool redirectStatus = response.StatusCode is 301 or 302 or 303 or 307 or 308;
                if (redirectStatus && response.Headers.TryGetValue("Location", out string? location)) {
                    if (++redirects > _policy.MaxRedirects)
                        throw new HtmlScriptRuntimeException("Resource redirect budget exceeded.");
                    Uri next = new(currentUrl, location);
                    CheckOrigin(next);
                    if (location.Contains('#')) responseFragment = next.Fragment;
                    if ((response.StatusCode is 301 or 302 && method == "POST") ||
                        (response.StatusCode == 303 && method is not ("GET" or "HEAD"))) {
                        method = "GET";
                        body = null;
                    }
                    _diagnostics.Record(HtmlRuntimeEventKind.Redirect, "document-redirect", "followed", started,
                        url: next, method: method, statusCode: response.StatusCode, redirectCount: redirects);
                    currentUrl = new Uri(HtmlRuntimeResourcePolicy.Key(next));
                    continue;
                }
                if (index != replay.Hops.Count - 1)
                    throw new HtmlScriptRuntimeException("The navigation replay contains responses after the document request completed.");
                Uri finalUrl = new(HtmlRuntimeResourcePolicy.Key(response.FinalUrl) + responseFragment);
                var result = new HtmlRuntimeResource(originalUrl, method == "HEAD" ? Array.Empty<byte>() : response.Buffer,
                    response.ContentType, response.StatusCode, finalUrl, redirects, response.Headers, response.StatusText);
                lock (_sync) _loaded[HtmlRuntimeResourcePolicy.Key(originalUrl)] = result;
                _diagnostics.Record(HtmlRuntimeEventKind.Resource, "document-navigation", "success", started,
                    DateTimeOffset.UtcNow - started, url: originalUrl, method: request.Method, statusCode: result.StatusCode,
                    byteCount: result.Length, redirectCount: redirects, decision: "replayed");
                return result;
            }
            throw new HtmlScriptRuntimeException("The navigation replay ended before the document request completed.");
        } catch (OperationCanceledException) when (deadline.IsCancellationRequested && !token.IsCancellationRequested && !_lifetime.IsCancellationRequested) {
            throw new HtmlScriptRuntimeException("The resource load exceeded its deadline.");
        } finally {
            if (admitted) _concurrency.Release();
            _budget.EndOperation();
        }
    }

    private async Task PreflightAsync(Uri url, string method, Dictionary<string, string> headers, CancellationToken token) {
        string[] unsafeHeaders = HtmlRuntimeCorsPolicy.UnsafeHeaders(headers);
        if (method is "GET" or "HEAD" or "POST" && unsafeHeaders.Length == 0) return;
        var preflightHeaders = new Dictionary<string, string> { ["Origin"] = _documentOrigin, ["Access-Control-Request-Method"] = method };
        if (unsafeHeaders.Length != 0) preflightHeaders["Access-Control-Request-Headers"] = string.Join(",", unsafeHeaders);
        var response = await SendAsync(url, "OPTIONS", preflightHeaders, null, false, token).ConfigureAwait(false);
        HtmlRuntimeCorsPolicy.CheckPreflight(response, _documentOrigin, method, unsafeHeaders);
    }

    private async Task<HtmlRuntimeResource> SendAsync(Uri url, string method, IReadOnlyDictionary<string, string> headers,
        byte[]? body, bool allowUrlSupply, CancellationToken token) {
        CheckOrigin(url);
        if (Interlocked.Increment(ref _budget.Requests) > _policy.MaxRequests) throw new HtmlScriptRuntimeException("Resource request budget exceeded.");
        if (allowUrlSupply && _supplied.TryGetValue(HtmlRuntimeResourcePolicy.Key(url), out var supplied) && method is "GET" or "HEAD" && body == null) {
            _diagnostics.Record(HtmlRuntimeEventKind.Policy, "resource-source", "allowed", DateTimeOffset.UtcNow,
                url: url, method: method, decision: "supplied");
            CheckOrigin(supplied.FinalUrl);
            if (supplied.RedirectCount > _policy.MaxRedirects) throw new HtmlScriptRuntimeException("Resource redirect budget exceeded.");
            if (Interlocked.Add(ref _budget.Requests, supplied.RedirectCount) > _policy.MaxRequests) throw new HtmlScriptRuntimeException("Resource request budget exceeded.");
            ReserveBytes(method == "HEAD" ? 0 : supplied.Length);
            var suppliedHeaders = new Dictionary<string, string>(supplied.Headers, StringComparer.OrdinalIgnoreCase) { ["Content-Type"] = supplied.ContentType };
            return new HtmlRuntimeResource(url, method == "HEAD" ? Array.Empty<byte>() : supplied.Buffer, supplied.ContentType, supplied.StatusCode, new Uri(HtmlRuntimeResourcePolicy.Key(supplied.FinalUrl)), supplied.RedirectCount, suppliedHeaders, supplied.StatusText);
        }
        if (!_policy.AllowNetwork) {
            bool replayable = method == "GET" && body == null && headers.Count == 0;
            if (replayable) _diagnostics.RecordMissingResource(url);
            _diagnostics.Record(HtmlRuntimeEventKind.Policy, "network-access", "blocked", DateTimeOffset.UtcNow,
                url: url, method: method, decision: replayable ? "network-disabled-replayable-get" : "network-disabled");
            throw new HtmlScriptRuntimeException(MissingResourceMessage);
        }
        _diagnostics.Record(HtmlRuntimeEventKind.Policy, "network-access", "allowed", DateTimeOffset.UtcNow,
            url: url, method: method, decision: "network-enabled");
        ReserveRequestBytes(body);
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

    private HtmlRuntimeFetchDiscovery NextFetchOccurrence(HtmlRuntimeFetchRequest request) {
        lock (_sync) {
            _budget.FetchOccurrences.TryGetValue(request.Identity, out int occurrence);
            occurrence++;
            _budget.FetchOccurrences[request.Identity] = occurrence;
            return new HtmlRuntimeFetchDiscovery(request, occurrence);
        }
    }

    private HtmlRuntimeNavigationDiscovery NextNavigationOccurrence(HtmlRuntimeNavigationRequest request) {
        lock (_sync) {
            _budget.NavigationOccurrences.TryGetValue(request.Identity, out int occurrence);
            occurrence++;
            _budget.NavigationOccurrences[request.Identity] = occurrence;
            return new HtmlRuntimeNavigationDiscovery(request, occurrence);
        }
    }

    private void ReserveExactRequest(byte[]? body) {
        if (Interlocked.Increment(ref _budget.Requests) > _policy.MaxRequests)
            throw new HtmlScriptRuntimeException("Resource request budget exceeded.");
        ReserveRequestBytes(body);
    }

    private void ReserveRequestBytes(byte[]? body) {
        if (body == null) return;
        lock (_sync) {
            if (body.LongLength > _policy.MaxTotalRequestBytes - _budget.SentBytes)
                throw new HtmlScriptRuntimeException("Total fetch request body byte budget exceeded.");
            _budget.SentBytes += body.LongLength;
        }
    }

    private void CheckOrigin(Uri url) {
        if (!_origins.Contains(HtmlRuntimeResourcePolicy.Origin(url))) {
            _diagnostics.Record(HtmlRuntimeEventKind.Policy, "origin", "blocked", DateTimeOffset.UtcNow,
                url: url, decision: "origin-not-allowed");
            throw new HtmlScriptRuntimeException("The resource origin is not allowed.");
        }
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
