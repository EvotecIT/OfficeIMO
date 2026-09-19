using System.Net;
using System.Net.Http.Headers;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Html.Runtime;

// Host-side acquisition for the future isolated profile. The worker receives
// immutable bytes and has no network route; this broker never runs in it.
internal sealed class HtmlPublicResourceBroker {
    private readonly HashSet<string> _staticAllowedHosts;
    private readonly HashSet<string> _allowedHosts;
    private readonly Func<string, CancellationToken, Task<IPAddress[]>> _resolveAddresses;
    private readonly Func<IPAddress, int, CancellationToken, ValueTask<Stream>> _connect;
    private readonly int _maxRequests;
    private readonly long _maxResourceBytes;
    private readonly long _maxTotalBytes;
    private readonly long _maxRequestBytes;
    private readonly long _maxTotalRequestBytes;
    private readonly TimeSpan _timeout;
    private readonly int _maxRedirects;
    private readonly object _budgetSync = new();
    private int _requests;
    private long _bytes;
    private long _requestBytes;

    internal HtmlPublicResourceBroker(IEnumerable<string> allowedHosts) : this(
        allowedHosts, (host, token) => Dns.GetHostAddressesAsync(host, token), ConnectToAddressAsync) { }

    internal HtmlPublicResourceBroker(IEnumerable<string> allowedHosts, int maxRequests,
        long maxResourceBytes, long maxTotalBytes, TimeSpan? timeout = null, int maxRedirects = 5,
        long maxRequestBytes = 1024 * 1024, long maxTotalRequestBytes = 16 * 1024 * 1024,
        IEnumerable<string>? dynamicHosts = null) : this(
        allowedHosts, (host, token) => Dns.GetHostAddressesAsync(host, token), ConnectToAddressAsync,
        maxRequests, maxResourceBytes, maxTotalBytes, timeout, maxRedirects, maxRequestBytes, maxTotalRequestBytes,
        dynamicHosts) { }

    internal HtmlPublicResourceBroker(IEnumerable<string> allowedHosts,
        Func<string, CancellationToken, Task<IPAddress[]>> resolveAddresses,
        Func<IPAddress, int, CancellationToken, ValueTask<Stream>> connect,
        int maxRequests = 32, long maxResourceBytes = 4 * 1024 * 1024,
        long maxTotalBytes = 16 * 1024 * 1024, TimeSpan? timeout = null, int maxRedirects = 5,
        long maxRequestBytes = 1024 * 1024, long maxTotalRequestBytes = 16 * 1024 * 1024,
        IEnumerable<string>? dynamicHosts = null) {
        ArgumentNullException.ThrowIfNull(allowedHosts);
        _resolveAddresses = resolveAddresses ?? throw new ArgumentNullException(nameof(resolveAddresses));
        _connect = connect ?? throw new ArgumentNullException(nameof(connect));
        if (maxRequests <= 0) throw new ArgumentOutOfRangeException(nameof(maxRequests));
        if (maxResourceBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxResourceBytes));
        if (maxTotalBytes < maxResourceBytes) throw new ArgumentOutOfRangeException(nameof(maxTotalBytes));
        if (maxRequestBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxRequestBytes));
        if (maxTotalRequestBytes < maxRequestBytes) throw new ArgumentOutOfRangeException(nameof(maxTotalRequestBytes));
        TimeSpan effectiveTimeout = timeout ?? TimeSpan.FromSeconds(20);
        if (effectiveTimeout <= TimeSpan.Zero || effectiveTimeout > TimeSpan.FromSeconds(20))
            throw new ArgumentOutOfRangeException(nameof(timeout));
        if (maxRedirects is < 0 or > 5) throw new ArgumentOutOfRangeException(nameof(maxRedirects));
        _maxRequests = maxRequests;
        _maxResourceBytes = maxResourceBytes;
        _maxTotalBytes = maxTotalBytes;
        _maxRequestBytes = maxRequestBytes;
        _maxTotalRequestBytes = maxTotalRequestBytes;
        _timeout = effectiveTimeout;
        _maxRedirects = maxRedirects;
        _staticAllowedHosts = new HashSet<string>(allowedHosts.Select(ValidateHost), StringComparer.OrdinalIgnoreCase);
        _allowedHosts = new HashSet<string>(_staticAllowedHosts, StringComparer.OrdinalIgnoreCase);
        if (dynamicHosts != null) _allowedHosts.UnionWith(dynamicHosts.Select(ValidateHost));
        if (_staticAllowedHosts.Count == 0 || _allowedHosts.Count > 16)
            throw new ArgumentException("The pilot requires one to sixteen explicitly allowed hosts.", nameof(allowedHosts));
    }

    internal bool AllowsHost(Uri url) => _staticAllowedHosts.Contains(ValidateHost(ValidateUrl(url).IdnHost));
    internal Uri[] AllowedOrigins => _allowedHosts.SelectMany(host => new[] {
        new Uri("http://" + host + "/"), new Uri("https://" + host + "/")
    }).ToArray();

    internal Task<HtmlPublicResourceResult> FetchAsync(Uri requestedUrl, CancellationToken cancellationToken = default) =>
        FetchStaticAsync(requestedUrl, cancellationToken);

    internal Task<HtmlPublicResourceResult> FetchAsync(HtmlRuntimeFetchDiscovery discovery, Uri documentUrl,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(discovery);
        ValidateDynamicRequest(discovery.Request);
        ValidateUrl(documentUrl);
        return FetchDynamicAsync(discovery, documentUrl, cancellationToken);
    }

    private async Task<HtmlPublicResourceResult> FetchDynamicAsync(HtmlRuntimeFetchDiscovery discovery, Uri documentUrl,
        CancellationToken cancellationToken) {
        HtmlRuntimeFetchRequest original = discovery.Request;
        Uri current = ValidateUrl(original.Url);
        string documentOrigin = HtmlRuntimeResourcePolicy.Origin(documentUrl);
        string requestOrigin = HtmlRuntimeResourcePolicy.Origin(current);
        bool cors = requestOrigin != documentOrigin;
        if (cors && original.Mode == "same-origin")
            throw new HtmlScriptRuntimeException("Cross-origin fetch is forbidden in same-origin mode.");
        string method = original.Method;
        byte[]? body = original.Buffer;
        var headers = new Dictionary<string, string>(original.Headers, StringComparer.OrdinalIgnoreCase);
        using var deadline = new CancellationTokenSource(_timeout);
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, deadline.Token);
        var redirects = new List<HtmlPublicRedirect>();
        var hops = new List<HtmlRuntimeFetchHop>();
        var exchanges = new List<HtmlPublicDynamicExchange>();
        DateTimeOffset fetchedAt = DateTimeOffset.UtcNow;
        IPAddress? finalAddress = null;
        for (int index = 0; ; index++) {
            operation.Token.ThrowIfCancellationRequested();
            HtmlRuntimeResource? preflight = null;
            if (cors && HtmlRuntimeCorsPolicy.RequiresPreflight(method, headers)) {
                string[] unsafeHeaders = HtmlRuntimeCorsPolicy.UnsafeHeaders(headers);
                var preflightHeaders = new Dictionary<string, string> {
                    ["Origin"] = documentOrigin, ["Access-Control-Request-Method"] = method
                };
                if (unsafeHeaders.Length != 0)
                    preflightHeaders["Access-Control-Request-Headers"] = string.Join(",", unsafeHeaders);
                var preflightResult = await SendDirectAsync(current, "OPTIONS", preflightHeaders, null, operation.Token).ConfigureAwait(false);
                preflight = preflightResult.Response;
                exchanges.Add(new HtmlPublicDynamicExchange(current, "OPTIONS", 0, null, preflight.StatusCode,
                    preflight.Length, Convert.ToHexString(SHA256.HashData(preflight.Buffer)).ToLowerInvariant(), preflightResult.Address));
                HtmlRuntimeCorsPolicy.CheckPreflight(preflight, documentOrigin, method, unsafeHeaders);
            }
            var outgoing = new Dictionary<string, string>(headers, StringComparer.OrdinalIgnoreCase);
            if (cors || method is not ("GET" or "HEAD")) outgoing["Origin"] = documentOrigin;
            (HtmlRuntimeResource response, IPAddress address) = await SendDirectAsync(current, method, outgoing, body, operation.Token).ConfigureAwait(false);
            finalAddress = address;
            exchanges.Add(new HtmlPublicDynamicExchange(current, method, body?.LongLength ?? 0,
                body == null ? null : Convert.ToHexString(SHA256.HashData(body)).ToLowerInvariant(), response.StatusCode,
                response.Length, Convert.ToHexString(SHA256.HashData(response.Buffer)).ToLowerInvariant(), address));
            hops.Add(new HtmlRuntimeFetchHop(response, preflight));
            if (cors) HtmlRuntimeCorsPolicy.Check(response.Headers, documentOrigin);
            if (response.StatusCode is 301 or 302 or 303 or 307 or 308 &&
                response.Headers.TryGetValue("Location", out string? location)) {
                if (original.Redirect == "error") throw new HtmlScriptRuntimeException("Fetch redirect mode forbids redirects.");
                if (index >= _maxRedirects) throw new HtmlScriptRuntimeException("The public resource redirect limit was exceeded.");
                Uri next = ValidateRedirect(current, ResolveRedirect(current, new Uri(location, UriKind.RelativeOrAbsolute)),
                    dynamic: true);
                if (HtmlRuntimeResourcePolicy.Origin(next) != requestOrigin)
                    throw new HtmlScriptRuntimeException("Dynamic redirects between distinct origins are outside the isolated profile.");
                redirects.Add(new HtmlPublicRedirect(current, next, response.StatusCode, address));
                if ((response.StatusCode is 301 or 302 && method == "POST") ||
                    (response.StatusCode == 303 && method is not ("GET" or "HEAD"))) {
                    method = "GET";
                    body = null;
                    foreach (string name in new[] { "Content-Encoding", "Content-Language", "Content-Location", "Content-Type" })
                        headers.Remove(name);
                }
                current = next;
                continue;
            }
            var resource = new HtmlRuntimeResource(original.Url, method == "HEAD" ? Array.Empty<byte>() : response.Buffer,
                response.ContentType, response.StatusCode, current, redirects.Count, response.Headers, response.StatusText);
            string digest = Convert.ToHexString(SHA256.HashData(resource.Buffer)).ToLowerInvariant();
            return new HtmlPublicResourceResult(resource, Array.AsReadOnly(redirects.ToArray()), fetchedAt,
                finalAddress, digest, discovery, Array.AsReadOnly(hops.ToArray()), Array.AsReadOnly(exchanges.ToArray()));
        }
    }

    private async Task<(HtmlRuntimeResource Response, IPAddress Address)> SendDirectAsync(Uri url, string method,
        IReadOnlyDictionary<string, string> headers, byte[]? body, CancellationToken token) {
        ReserveRequest();
        if (body?.LongLength > _maxRequestBytes)
            throw new HtmlScriptRuntimeException("The dynamic request body exceeds its byte budget.");
        if (body != null) lock (_budgetSync) {
            if (body.LongLength > _maxTotalRequestBytes - _requestBytes)
                throw new HtmlScriptRuntimeException("The dynamic request bodies exceed their cumulative byte budget.");
            _requestBytes += body.LongLength;
        }
        IPAddress? connectedAddress = null;
        using var handler = new SocketsHttpHandler {
            AllowAutoRedirect = false, UseCookies = false, UseProxy = false,
            AutomaticDecompression = DecompressionMethods.None,
            ConnectCallback = async (context, callbackToken) => {
                string host = ValidateHost(context.DnsEndPoint.Host);
                if (!_allowedHosts.Contains(host)) throw new HtmlScriptRuntimeException("The public resource host is not allowed.");
                IPAddress address = SelectPublicAddress(await _resolveAddresses(host, callbackToken).ConfigureAwait(false));
                Stream stream = await _connect(address, context.DnsEndPoint.Port, callbackToken).ConfigureAwait(false);
                connectedAddress = address;
                return stream;
            }
        };
        using var client = new HttpClient(handler) { Timeout = Timeout.InfiniteTimeSpan };
        using var request = new HttpRequestMessage(new HttpMethod(method), WithoutFragment(url));
        if (body != null) request.Content = new ByteArrayContent(body);
        foreach (var header in headers) {
            if (!request.Headers.TryAddWithoutValidation(header.Key, header.Value)) {
                request.Content ??= new ByteArrayContent(Array.Empty<byte>());
                request.Content.Headers.TryAddWithoutValidation(header.Key, header.Value);
            }
        }
        request.Headers.AcceptEncoding.Add(new StringWithQualityHeaderValue("identity"));
        request.Headers.UserAgent.ParseAdd("OfficeIMO-HTML-Pilot/1.0");
        using HttpResponseMessage response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, token).ConfigureAwait(false);
        if (connectedAddress == null) throw new HtmlScriptRuntimeException("The public resource address was not verified.");
        if (response.Content.Headers.ContentEncoding.Any(value => !value.Equals("identity", StringComparison.OrdinalIgnoreCase)))
            throw new HtmlScriptRuntimeException("Compressed public responses are outside the pilot transport profile.");
        if (method != "HEAD" && response.Content.Headers.ContentLength > _maxResourceBytes)
            throw new HtmlScriptRuntimeException("The public resource exceeds its byte budget.");
        byte[] bytes = method == "HEAD" ? Array.Empty<byte>() : await ReadBoundedAsync(response.Content, token).ConfigureAwait(false);
        var responseHeaders = response.Headers.Concat(response.Content.Headers)
            .ToDictionary(header => header.Key, header => string.Join(", ", header.Value), StringComparer.OrdinalIgnoreCase);
        return (new HtmlRuntimeResource(url, bytes, response.Content.Headers.ContentType?.ToString() ?? "application/octet-stream",
            (int)response.StatusCode, headers: responseHeaders, statusText: response.ReasonPhrase ?? string.Empty), connectedAddress);
    }

    private async Task<HtmlPublicResourceResult> FetchStaticAsync(Uri requestedUrl, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(requestedUrl);
        Uri current = ValidateUrl(requestedUrl);
        if (!AllowsHost(current))
            throw new HtmlScriptRuntimeException("The public resource host is not allowed.");
        using var deadline = new CancellationTokenSource(_timeout);
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, deadline.Token);
        var redirects = new List<HtmlPublicRedirect>();
        DateTimeOffset fetchedAt = DateTimeOffset.UtcNow;
        for (int hop = 0; ; hop++) {
            operation.Token.ThrowIfCancellationRequested();
            (HtmlRuntimeResource response, IPAddress connectedAddress) = await SendDirectAsync(current, "GET",
                new Dictionary<string, string>(), null, operation.Token).ConfigureAwait(false);
            int status = response.StatusCode;
            if (status is 301 or 302 or 303 or 307 or 308) {
                if (hop >= _maxRedirects || !response.Headers.TryGetValue("Location", out string? location))
                    throw new HtmlScriptRuntimeException("The public resource redirect limit or location was invalid.");
                Uri next = ValidateRedirect(current, ResolveRedirect(current, new Uri(location, UriKind.RelativeOrAbsolute)));
                redirects.Add(new HtmlPublicRedirect(current, next, status, connectedAddress));
                current = next;
                continue;
            }
            if (status is < 200 or >= 300)
                throw new HtmlScriptRuntimeException("The public resource returned HTTP " + status + ".");
            var resource = new HtmlRuntimeResource(requestedUrl, response.Buffer, response.ContentType, status,
                current, redirects.Count, response.Headers, response.StatusText);
            string digest = Convert.ToHexString(SHA256.HashData(response.Buffer)).ToLowerInvariant();
            return new HtmlPublicResourceResult(resource, Array.AsReadOnly(redirects.ToArray()), fetchedAt,
                connectedAddress, digest);
        }
    }

    internal static void ValidateDynamicRequest(HtmlRuntimeFetchRequest request) {
        ArgumentNullException.ThrowIfNull(request);
        ValidateUrl(request.Url);
        if (request.Headers.Keys.Any(name => name.Equals("Authorization", StringComparison.OrdinalIgnoreCase)
                || name.Equals("Proxy-Authorization", StringComparison.OrdinalIgnoreCase)))
            throw new HtmlScriptRuntimeException("Credential-bearing dynamic request headers are forbidden in the isolated profile.");
        if (request.Credentials is not ("omit" or "same-origin"))
            throw new HtmlScriptRuntimeException("Unsupported dynamic request credentials mode.");
    }

    internal static Uri ValidateUrl(Uri url) {
        HtmlRuntimeResourcePolicy.ValidateUrl(url);
        int standardPort = url.Scheme == Uri.UriSchemeHttps ? 443 : 80;
        if (url.Port != standardPort || url.HostNameType is not (UriHostNameType.Dns or UriHostNameType.IPv4))
            throw new ArgumentException("The public pilot accepts DNS names or IPv4 literals on standard HTTP(S) ports only.", nameof(url));
        return url;
    }

    internal Uri ValidateRedirect(Uri from, Uri target, bool dynamic = false) {
        target = ValidateUrl(target);
        if (from.Scheme == Uri.UriSchemeHttps && target.Scheme != Uri.UriSchemeHttps)
            throw new HtmlScriptRuntimeException("HTTPS to HTTP redirects are forbidden in the public pilot.");
        if (!(dynamic ? _allowedHosts : _staticAllowedHosts).Contains(ValidateHost(target.IdnHost)))
            throw new HtmlScriptRuntimeException("The public resource redirect host is not allowed.");
        return target;
    }

    internal static Uri ResolveRedirect(Uri from, Uri location) {
        Uri target = new(from, location);
        return location.OriginalString.Contains('#') ? target
            : new UriBuilder(target) { Fragment = from.Fragment.TrimStart('#') }.Uri;
    }

    internal static IPAddress SelectPublicAddress(IReadOnlyList<IPAddress> addresses) {
        ArgumentNullException.ThrowIfNull(addresses);
        IPAddress[] ipv4 = addresses.Where(address => address.AddressFamily == AddressFamily.InterNetwork).ToArray();
        if (ipv4.Length == 0 || ipv4.Any(address => !IsPublicIpv4(address)))
            throw new HtmlScriptRuntimeException("The public resource resolved to an unsupported or non-public address.");
        return ipv4[0];
    }

    internal static bool IsPublicIpv4(IPAddress address) {
        if (address.AddressFamily != AddressFamily.InterNetwork) return false;
        byte[] octets = address.GetAddressBytes();
        byte a = octets[0], b = octets[1], c = octets[2];
        if (a is 0 or 10 or 127 or >= 224) return false;
        if (a == 169 && b == 254 || a == 100 && b is >= 64 and <= 127 || a == 172 && b is >= 16 and <= 31)
            return false;
        if (a == 192 && (b == 0 && c is 0 or 2 || b == 88 && c == 99 || b == 168)) return false;
        if (a == 198 && (b is 18 or 19 || b == 51 && c == 100) || a == 203 && b == 0 && c == 113) return false;
        return true;
    }

    internal static string DecodeUtf8Html(HtmlRuntimeResource resource, bool allowXhtml = true) {
        string mediaType = resource.ContentType.Split(';', 2)[0].Trim();
        if (!mediaType.Equals("text/html", StringComparison.OrdinalIgnoreCase)
            && (!allowXhtml || !mediaType.Equals("application/xhtml+xml", StringComparison.OrdinalIgnoreCase)))
            throw new HtmlScriptRuntimeException("The public document is not HTML.");
        string[] parameters = resource.ContentType.Split(';');
        foreach (string parameter in parameters.Skip(1)) {
            string value = parameter.Trim();
            if (value.StartsWith("charset=", StringComparison.OrdinalIgnoreCase)
                && !value[8..].Trim(' ', '"').Equals("utf-8", StringComparison.OrdinalIgnoreCase))
                throw new HtmlScriptRuntimeException("The public pilot accepts UTF-8 HTML only.");
        }
        try { return new UTF8Encoding(false, true).GetString(resource.Buffer).TrimStart('\uFEFF'); }
        catch (DecoderFallbackException) { throw new HtmlScriptRuntimeException("The public HTML response is not valid UTF-8."); }
    }

    private static string ValidateHost(string host) {
        if (string.IsNullOrWhiteSpace(host) || host.Length > 253 || host.Contains('/') || host.Contains(':') || host.Contains('@'))
            throw new ArgumentException("The public resource host is invalid.", nameof(host));
        string canonical = host.TrimEnd('.').ToLowerInvariant();
        if (canonical.Length == 0 || canonical.Any(char.IsWhiteSpace)) throw new ArgumentException("The public resource host is invalid.", nameof(host));
        return canonical;
    }

    private static Uri WithoutFragment(Uri url) => new UriBuilder(url) { Fragment = string.Empty }.Uri;

    internal static async ValueTask<Stream> ConnectToAddressAsync(IPAddress address, int port, CancellationToken token) {
        var socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        try {
            await socket.ConnectAsync(address, port, token).ConfigureAwait(false);
            return new NetworkStream(socket, ownsSocket: true);
        } catch { socket.Dispose(); throw; }
    }

    private void ReserveRequest() {
        lock (_budgetSync) if (++_requests > _maxRequests) throw new HtmlScriptRuntimeException("The public acquisition request budget was exceeded.");
    }

    private async Task<byte[]> ReadBoundedAsync(HttpContent content, CancellationToken token) {
        await using Stream input = await content.ReadAsStreamAsync(token).ConfigureAwait(false);
        using var output = new MemoryStream();
        var buffer = new byte[16 * 1024];
        int read;
        while ((read = await input.ReadAsync(buffer, token).ConfigureAwait(false)) != 0) {
            lock (_budgetSync) {
                if (read > _maxResourceBytes - output.Length || read > _maxTotalBytes - _bytes)
                    throw new HtmlScriptRuntimeException("The public acquisition byte budget was exceeded.");
                _bytes += read;
            }
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }
}

internal sealed record HtmlPublicRedirect(Uri From, Uri To, int StatusCode, IPAddress ConnectedAddress);
internal sealed record HtmlPublicResourceResult(HtmlRuntimeResource Resource, IReadOnlyList<HtmlPublicRedirect> Redirects,
    DateTimeOffset FetchedAtUtc, IPAddress ConnectedAddress, string Sha256,
    HtmlRuntimeFetchDiscovery? DynamicRequest = null, IReadOnlyList<HtmlRuntimeFetchHop>? DynamicHops = null,
    IReadOnlyList<HtmlPublicDynamicExchange>? DynamicExchanges = null);

internal sealed record HtmlPublicDynamicExchange(Uri Url, string Method, long RequestBodyByteCount,
    string? RequestBodySha256, int StatusCode, long ResponseByteCount, string ResponseSha256, IPAddress ConnectedAddress);
