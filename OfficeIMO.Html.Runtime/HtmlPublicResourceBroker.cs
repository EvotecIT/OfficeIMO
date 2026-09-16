using System.Net;
using System.Net.Http.Headers;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Html.Runtime;

// Host-side acquisition for the future isolated profile. The worker receives
// immutable bytes and has no network route; this broker never runs in it.
internal sealed class HtmlPublicResourceBroker {
    private readonly HashSet<string> _allowedHosts;
    private readonly object _budgetSync = new();
    private int _requests;
    private long _bytes;

    internal HtmlPublicResourceBroker(IEnumerable<string> allowedHosts) {
        ArgumentNullException.ThrowIfNull(allowedHosts);
        _allowedHosts = new HashSet<string>(allowedHosts.Select(ValidateHost), StringComparer.OrdinalIgnoreCase);
        if (_allowedHosts.Count == 0 || _allowedHosts.Count > 16)
            throw new ArgumentException("The pilot requires one to sixteen explicitly allowed hosts.", nameof(allowedHosts));
    }

    internal async Task<HtmlPublicResourceResult> FetchAsync(Uri requestedUrl, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(requestedUrl);
        Uri current = ValidateUrl(requestedUrl);
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(20));
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, deadline.Token);
        var redirects = new List<HtmlPublicRedirect>();
        DateTimeOffset fetchedAt = DateTimeOffset.UtcNow;
        for (int hop = 0; ; hop++) {
            operation.Token.ThrowIfCancellationRequested();
            ReserveRequest();
            IPAddress? connectedAddress = null;
            using var handler = new SocketsHttpHandler {
                AllowAutoRedirect = false, UseCookies = false, UseProxy = false,
                AutomaticDecompression = DecompressionMethods.None,
                ConnectCallback = async (context, token) => {
                    string host = ValidateHost(context.DnsEndPoint.Host);
                    if (!_allowedHosts.Contains(host)) throw new HtmlScriptRuntimeException("The public resource host is not allowed.");
                    IPAddress address = SelectPublicAddress(await Dns.GetHostAddressesAsync(host, token).ConfigureAwait(false));
                    var socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                    try {
                        await socket.ConnectAsync(address, context.DnsEndPoint.Port, token).ConfigureAwait(false);
                        connectedAddress = address;
                        return new NetworkStream(socket, ownsSocket: true);
                    } catch { socket.Dispose(); throw; }
                }
            };
            using var client = new HttpClient(handler) { Timeout = Timeout.InfiniteTimeSpan };
            using var request = new HttpRequestMessage(HttpMethod.Get, WithoutFragment(current));
            request.Headers.AcceptEncoding.Add(new StringWithQualityHeaderValue("identity"));
            request.Headers.UserAgent.ParseAdd("OfficeIMO-HTML-Pilot/1.0");
            using HttpResponseMessage response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, operation.Token).ConfigureAwait(false);
            if (connectedAddress == null) throw new HtmlScriptRuntimeException("The public resource address was not verified.");
            int status = (int)response.StatusCode;
            if (status is 301 or 302 or 303 or 307 or 308) {
                if (hop >= 5 || response.Headers.Location == null)
                    throw new HtmlScriptRuntimeException("The public resource redirect limit or location was invalid.");
                Uri next = ValidateRedirect(current, ResolveRedirect(current, response.Headers.Location));
                redirects.Add(new HtmlPublicRedirect(current, next, status, connectedAddress));
                current = next;
                continue;
            }
            if (!response.IsSuccessStatusCode)
                throw new HtmlScriptRuntimeException("The public resource returned HTTP " + status + ".");
            if (response.Content.Headers.ContentEncoding.Any(value => !value.Equals("identity", StringComparison.OrdinalIgnoreCase)))
                throw new HtmlScriptRuntimeException("Compressed public responses are outside the pilot transport profile.");
            if (response.Content.Headers.ContentLength is > 4 * 1024 * 1024)
                throw new HtmlScriptRuntimeException("The public resource exceeds its byte budget.");
            byte[] bytes = await ReadBoundedAsync(response.Content, operation.Token).ConfigureAwait(false);
            string contentType = response.Content.Headers.ContentType?.ToString() ?? "application/octet-stream";
            var resource = new HtmlRuntimeResource(requestedUrl, bytes, contentType, status,
                current, redirects.Count, statusText: response.ReasonPhrase ?? string.Empty);
            string digest = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
            return new HtmlPublicResourceResult(resource, Array.AsReadOnly(redirects.ToArray()), fetchedAt, connectedAddress, digest);
        }
    }

    internal static Uri ValidateUrl(Uri url) {
        HtmlRuntimeResourcePolicy.ValidateUrl(url);
        int standardPort = url.Scheme == Uri.UriSchemeHttps ? 443 : 80;
        if (url.Port != standardPort || url.HostNameType is not (UriHostNameType.Dns or UriHostNameType.IPv4))
            throw new ArgumentException("The public pilot accepts DNS names or IPv4 literals on standard HTTP(S) ports only.", nameof(url));
        return url;
    }

    internal Uri ValidateRedirect(Uri from, Uri target) {
        target = ValidateUrl(target);
        if (from.Scheme == Uri.UriSchemeHttps && target.Scheme != Uri.UriSchemeHttps)
            throw new HtmlScriptRuntimeException("HTTPS to HTTP redirects are forbidden in the public pilot.");
        if (!_allowedHosts.Contains(ValidateHost(target.IdnHost)))
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

    internal static string DecodeUtf8Html(HtmlRuntimeResource resource) {
        string mediaType = resource.ContentType.Split(';', 2)[0].Trim();
        if (!mediaType.Equals("text/html", StringComparison.OrdinalIgnoreCase)
            && !mediaType.Equals("application/xhtml+xml", StringComparison.OrdinalIgnoreCase))
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

    private void ReserveRequest() {
        lock (_budgetSync) if (++_requests > 32) throw new HtmlScriptRuntimeException("The public acquisition request budget was exceeded.");
    }

    private async Task<byte[]> ReadBoundedAsync(HttpContent content, CancellationToken token) {
        await using Stream input = await content.ReadAsStreamAsync(token).ConfigureAwait(false);
        using var output = new MemoryStream();
        var buffer = new byte[16 * 1024];
        int read;
        while ((read = await input.ReadAsync(buffer, token).ConfigureAwait(false)) != 0) {
            lock (_budgetSync) {
                if (read > 4 * 1024 * 1024 - output.Length || read > 16 * 1024 * 1024 - _bytes)
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
    DateTimeOffset FetchedAtUtc, IPAddress ConnectedAddress, string Sha256);
