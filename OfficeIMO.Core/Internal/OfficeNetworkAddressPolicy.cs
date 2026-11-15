using System.Net;
using System.Net.Http;
using System.Net.Sockets;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#elif NETSTANDARD2_0
using System.Runtime.InteropServices;
#endif

namespace OfficeIMO.Core.Internal;

/// <summary>
/// Raised when a remote request would reach a loopback, private, link-local or other non-public address.
/// </summary>
internal sealed class OfficeNetworkAddressException : IOException {
    internal OfficeNetworkAddressException(string host)
        : base("The host '" + host + "' is or resolves to a loopback, private, link-local or other non-public network address. Set AllowPrivateNetworkAddresses only for trusted internal sources.") {
        Host = host;
    }

    internal string Host { get; }
}

/// <summary>
/// The one place OfficeIMO decides whether an HTTP(S) destination is on the public internet.
/// </summary>
/// <remarks>
/// Fetchers owned by OfficeIMO send through <see cref="GetClient"/> or <see cref="CreateHandler"/> and build each
/// request, including every redirect hop, with <see cref="CreateRequestAsync"/>. On .NET 8 and later the handler
/// resolves the host when it opens a connection, checks every address and connects to the checked address, so a
/// second DNS answer cannot move the request to an internal address. Older frameworks have no connect callback:
/// the host is resolved and checked before the request, plain HTTP is sent to the checked address with the original
/// Host header, and HTTPS connects by name, where certificate validation binds the connection to that name. Browser
/// runtimes cannot resolve names, so only IP literals and localhost names are checked there.
/// </remarks>
internal static class OfficeNetworkAddressPolicy {
    private static readonly Lazy<HttpClient> PublicClient = new(() => CreateClient(allowPrivateNetworkAddresses: false));
    private static readonly Lazy<HttpClient> PrivateClient = new(() => CreateClient(allowPrivateNetworkAddresses: true));

#if NET8_0_OR_GREATER
    [SupportedOSPlatformGuard("browser")]
    private static bool IsBrowser => OperatingSystem.IsBrowser();
#elif NETSTANDARD2_0
    private static bool IsBrowser => RuntimeInformation.IsOSPlatform(OSPlatform.Create("BROWSER"));
#else
    private static bool IsBrowser => false;
#endif

    /// <summary>Returns a pooled client with redirects and cookies disabled that enforces the address policy.</summary>
    internal static HttpClient GetClient(bool allowPrivateNetworkAddresses) =>
        allowPrivateNetworkAddresses ? PrivateClient.Value : PublicClient.Value;

    /// <summary>Creates a handler with redirects and cookies disabled that enforces the address policy.</summary>
    /// <remarks>
    /// While non-public addresses are refused, the handler connects directly; a proxy would otherwise be the
    /// address that is checked and connected to.
    /// </remarks>
    internal static HttpMessageHandler CreateHandler(bool allowPrivateNetworkAddresses, DecompressionMethods decompression) {
#if NET8_0_OR_GREATER
        if (SocketsHttpHandler.IsSupported) {
            var sockets = new SocketsHttpHandler {
                AllowAutoRedirect = false,
                UseCookies = false,
                AutomaticDecompression = decompression,
                PooledConnectionIdleTimeout = TimeSpan.FromMinutes(1),
                PooledConnectionLifetime = TimeSpan.FromMinutes(5)
            };
            if (!allowPrivateNetworkAddresses) {
                sockets.UseProxy = false;
                sockets.ConnectCallback = ConnectToPublicAddressAsync;
            }
            return sockets;
        }
#endif
        var handler = new HttpClientHandler { AllowAutoRedirect = false };
        if (!IsBrowser) {
            handler.UseCookies = false;
            handler.AutomaticDecompression = decompression;
            if (!allowPrivateNetworkAddresses) handler.UseProxy = false;
        }
        return handler;
    }

    /// <summary>
    /// Creates a request for one hop sent through a handler from <see cref="CreateHandler"/>, refusing a
    /// non-public destination unless private addresses are allowed.
    /// </summary>
    /// <remarks>Cookies and redirects must be resolved against <paramref name="uri"/>, not the request URI.</remarks>
    internal static async Task<HttpRequestMessage> CreateRequestAsync(
        HttpMethod method,
        Uri uri,
        bool allowPrivateNetworkAddresses,
        CancellationToken cancellationToken) {
        if (uri == null) throw new ArgumentNullException(nameof(uri));
        if (allowPrivateNetworkAddresses) return new HttpRequestMessage(method, uri);
        if (IsBrowser) {
            ThrowIfNonPublicHostName(uri);
            return new HttpRequestMessage(method, uri);
        }
        ThrowIfNonPublicLiteral(uri);
        if (ChecksAtConnection) return new HttpRequestMessage(method, uri);

        IPAddress[] addresses = await ResolvePublicAsync(HostOf(uri), cancellationToken).ConfigureAwait(false);
        return string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
            ? CreatePinnedRequest(method, uri, addresses[0])
            : new HttpRequestMessage(method, uri);
    }

    /// <summary>
    /// Screens a destination that a caller-supplied client will resolve and connect to: refuses non-public IP
    /// literals and localhost names. It cannot see the addresses such a client connects to.
    /// </summary>
    internal static void ThrowIfNonPublicHostName(Uri uri) {
        ThrowIfNonPublicLiteral(uri);
        string host = HostOf(uri);
        if (string.Equals(host, "localhost", StringComparison.OrdinalIgnoreCase)
            || host.EndsWith(".localhost", StringComparison.OrdinalIgnoreCase)) {
            throw new OfficeNetworkAddressException(host);
        }
    }

    /// <summary>Returns the policy rejection carried by an exception or one of its inner exceptions.</summary>
    internal static bool TryGetRejection(Exception? exception, out OfficeNetworkAddressException rejection) {
        for (Exception? current = exception; current != null; current = current.InnerException) {
            if (current is OfficeNetworkAddressException found) {
                rejection = found;
                return true;
            }
        }
        rejection = null!;
        return false;
    }

    /// <summary>Returns whether an address may be contacted without explicit permission.</summary>
    internal static bool IsPublic(IPAddress address) {
        if (address == null) throw new ArgumentNullException(nameof(address));
        if (address.IsIPv4MappedToIPv6) return IsPublic(address.MapToIPv4());
        byte[] b = address.GetAddressBytes();
        if (address.AddressFamily == AddressFamily.InterNetwork) return IsPublicIPv4(b, 0);
        if (address.AddressFamily != AddressFamily.InterNetworkV6 || b.Length != 16) return false;
        if (address.IsIPv6LinkLocal || address.IsIPv6SiteLocal || address.IsIPv6Multicast) return false;
        if ((b[0] & 0xFE) == 0xFC) return false;                                           // fc00::/7 unique local
        if (AllZero(b, 0, 12)) return false;                                               // ::, ::1 and deprecated ::a.b.c.d
        if (AllZero(b, 0, 8) && b[8] == 0xFF && b[9] == 0xFF && b[10] == 0 && b[11] == 0) {
            return IsPublicIPv4(b, 12);                                                    // ::ffff:0:a.b.c.d translated
        }
        if (b[0] == 0x01 && AllZero(b, 1, 7)) return false;                                // 100::/64 discard only
        if (b[0] == 0x20 && b[1] == 0x01 && b[2] == 0x0D && b[3] == 0xB8) return false;   // 2001:db8::/32 documentation
        if (b[0] == 0x20 && b[1] == 0x01 && b[2] < 0x02) return false;                     // 2001::/23 protocol assignments, Teredo
        if (b[0] == 0x20 && b[1] == 0x02) return IsPublicIPv4(b, 2);                      // 2002::/16 6to4 embeds IPv4
        if (b[0] == 0x00 && b[1] == 0x64 && b[2] == 0xFF && b[3] == 0x9B) {
            if (b[4] == 0x00 && b[5] == 0x01) return false;                                // 64:ff9b:1::/48 local-use translation
            if (AllZero(b, 4, 8)) return IsPublicIPv4(b, 12);                             // 64:ff9b::/96 NAT64 embeds IPv4
        }
        return true;
    }

    private static bool IsPublicIPv4(byte[] b, int offset) {
        byte first = b[offset], second = b[offset + 1], third = b[offset + 2], fourth = b[offset + 3];
        if (first == 0 || first == 10 || first == 127 || first >= 224) return false;      // this network, private, loopback, multicast/reserved
        if (first == 100 && (second & 0xC0) == 64) return false;                           // 100.64.0.0/10 shared address space
        if (first == 169 && second == 254) return false;                                   // link-local, including cloud metadata
        if (first == 172 && (second & 0xF0) == 16) return false;                           // 172.16.0.0/12
        if (first == 192 && second == 168) return false;                                   // 192.168.0.0/16
        if (first == 192 && second == 0 && third == 0) return fourth == 9 || fourth == 10; // 192.0.0.0/24 except global anycast
        if (first == 192 && second == 0 && third == 2) return false;                       // TEST-NET-1
        if (first == 192 && second == 88 && third == 99) return false;                     // deprecated 6to4 relay anycast
        if (first == 198 && (second & 0xFE) == 18) return false;                           // 198.18.0.0/15 benchmarking
        if (first == 198 && second == 51 && third == 100) return false;                    // TEST-NET-2
        if (first == 203 && second == 0 && third == 113) return false;                     // TEST-NET-3
        return true;
    }

    private static bool AllZero(byte[] bytes, int offset, int count) {
        for (int index = offset; index < offset + count; index++) {
            if (bytes[index] != 0) return false;
        }
        return true;
    }

    private static void ThrowIfNonPublicLiteral(Uri uri) {
        string host = HostOf(uri);
        if (IPAddress.TryParse(host, out IPAddress? literal) && !IsPublic(literal)) {
            throw new OfficeNetworkAddressException(host);
        }
    }

    private static string HostOf(Uri uri) {
        if (!uri.IsAbsoluteUri) throw new ArgumentException("The URI must be absolute.", nameof(uri));
        return uri.HostNameType == UriHostNameType.IPv6 ? uri.Host.Trim('[', ']') : uri.IdnHost.TrimEnd('.');
    }

    private static async Task<IPAddress[]> ResolvePublicAsync(string host, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        IPAddress[] addresses;
        if (IPAddress.TryParse(host, out IPAddress? literal)) {
            addresses = new[] { literal };
        } else {
#if NET6_0_OR_GREATER
            addresses = await Dns.GetHostAddressesAsync(host, cancellationToken).ConfigureAwait(false);
#else
            Task<IPAddress[]> resolution = Dns.GetHostAddressesAsync(host);
            Task completed = await Task.WhenAny(resolution, Task.Delay(Timeout.Infinite, cancellationToken)).ConfigureAwait(false);
            if (completed != resolution) {
                _ = resolution.ContinueWith(static task => _ = task.Exception, TaskContinuationOptions.OnlyOnFaulted);
                cancellationToken.ThrowIfCancellationRequested();
            }
            addresses = await resolution.ConfigureAwait(false);
#endif
        }
        if (addresses.Length == 0) throw new HttpRequestException("The host '" + host + "' did not resolve to an address.");
        foreach (IPAddress address in addresses) {
            if (!IsPublic(address)) throw new OfficeNetworkAddressException(host);
        }
        return addresses;
    }

    /// <summary>Addresses a plain HTTP request to a checked address while naming the original host.</summary>
    internal static HttpRequestMessage CreatePinnedRequest(HttpMethod method, Uri uri, IPAddress address) {
        string literal = address.AddressFamily == AddressFamily.InterNetworkV6 ? "[" + address + "]" : address.ToString();
        var request = new HttpRequestMessage(method, new UriBuilder(uri) { Host = literal }.Uri);
        string host = uri.HostNameType == UriHostNameType.IPv6 ? uri.Host : uri.IdnHost;
        request.Headers.Host = uri.IsDefaultPort ? host : host + ":" + uri.Port.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return request;
    }

    private static HttpClient CreateClient(bool allowPrivateNetworkAddresses) =>
        new(CreateHandler(allowPrivateNetworkAddresses, DecompressionMethods.GZip | DecompressionMethods.Deflate), disposeHandler: true) {
            Timeout = Timeout.InfiniteTimeSpan
        };

#if NET8_0_OR_GREATER
    private static bool ChecksAtConnection => SocketsHttpHandler.IsSupported;

    private static async ValueTask<Stream> ConnectToPublicAddressAsync(SocketsHttpConnectionContext context, CancellationToken cancellationToken) {
        IPAddress[] addresses = await ResolvePublicAsync(context.DnsEndPoint.Host, cancellationToken).ConfigureAwait(false);
        Exception? lastFailure = null;
        foreach (IPAddress address in addresses) {
            var socket = new Socket(address.AddressFamily, SocketType.Stream, ProtocolType.Tcp) { NoDelay = true };
            try {
                await socket.ConnectAsync(new IPEndPoint(address, context.DnsEndPoint.Port), cancellationToken).ConfigureAwait(false);
                return new NetworkStream(socket, ownsSocket: true);
            } catch (SocketException exception) {
                socket.Dispose();
                lastFailure = exception;
            } catch {
                socket.Dispose();
                throw;
            }
        }
        throw new HttpRequestException("No checked address of '" + context.DnsEndPoint.Host + "' accepted a connection.", lastFailure);
    }
#else
    private static bool ChecksAtConnection => false;
#endif
}
