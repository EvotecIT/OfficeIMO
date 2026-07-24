using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Drawing;

/// <summary>
/// Retrieves bounded image content for OfficeIMO packages that explicitly opt into remote I/O.
/// </summary>
public static class OfficeRemoteImageLoader {
    private const int BufferSize = 81920;
#if NET8_0_OR_GREATER
    private static readonly HttpClient PublicClient = CreatePinnedClient(
        allowPrivateNetworkAddresses: false);
    private static readonly HttpClient PrivateClient = CreatePinnedClient(
        allowPrivateNetworkAddresses: true);
#else
    private static readonly HttpClient LegacyClient = CreateClient();
#endif

    /// <summary>
    /// Asynchronously retrieves an image from an HTTP or HTTPS URL.
    /// </summary>
    public static Task<OfficeRemoteImage> LoadAsync(
        string url,
        OfficeRemoteImageLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(url)) {
            throw new ArgumentException("Image URL cannot be null or whitespace.", nameof(url));
        }

        if (!Uri.TryCreate(url, UriKind.Absolute, out Uri? uri)) {
            throw new ArgumentException("Image URL must be an absolute HTTP or HTTPS URL.", nameof(url));
        }

        return LoadAsync(uri, options, cancellationToken);
    }

    /// <summary>
    /// Asynchronously retrieves an image from an HTTP or HTTPS URI.
    /// </summary>
    public static async Task<OfficeRemoteImage> LoadAsync(
        Uri uri,
        OfficeRemoteImageLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (uri == null) throw new ArgumentNullException(nameof(uri));
        ValidateHttpUri(uri, nameof(uri));

        OfficeRemoteImageLoadOptions.Snapshot snapshot =
            (options ?? new OfficeRemoteImageLoadOptions()).CreateSnapshot();
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        if (snapshot.Timeout != System.Threading.Timeout.InfiniteTimeSpan) {
            timeout.CancelAfter(snapshot.Timeout);
        }

        Uri current = uri;
        for (int redirectCount = 0; ; redirectCount++) {
            timeout.Token.ThrowIfCancellationRequested();
            ValidateDestinationPolicy(current, snapshot);
#if NET8_0_OR_GREATER
            HttpClient client = snapshot.AllowPrivateNetworkAddresses
                ? PrivateClient
                : PublicClient;
            using var request = new HttpRequestMessage(HttpMethod.Get,
                current);
#else
            IPAddress[] addresses = await ResolveDestinationAsync(current,
                snapshot.AllowPrivateNetworkAddresses, timeout.Token)
                .ConfigureAwait(false);
            HttpClient client = LegacyClient;
            using HttpRequestMessage request = CreatePinnedRequest(current,
                addresses[0]);
#endif
            using HttpResponseMessage response = await client.SendAsync(
                request,
                HttpCompletionOption.ResponseHeadersRead,
                timeout.Token).ConfigureAwait(false);

            if (IsRedirect(response.StatusCode)) {
                if (redirectCount >= snapshot.MaximumRedirects) {
                    throw new InvalidDataException($"The remote image exceeded the {snapshot.MaximumRedirects}-redirect limit.");
                }

                Uri next = ResolveRedirect(current, response.Headers.Location);
                current = next;
                continue;
            }

            response.EnsureSuccessStatusCode();
            string contentType = NormalizeImageContentType(response.Content.Headers.ContentType?.MediaType);
            long? contentLength = response.Content.Headers.ContentLength;
            if (contentLength.HasValue && contentLength.Value > snapshot.MaximumBytes) {
                throw new InvalidDataException($"The remote image exceeds the {snapshot.MaximumBytes}-byte limit.");
            }

            using Stream input = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            byte[] bytes = await ReadBoundedAsync(input, snapshot.MaximumBytes, timeout.Token).ConfigureAwait(false);
            string fileName = Path.GetFileName(current.LocalPath);
            if (string.IsNullOrWhiteSpace(fileName)) {
                fileName = "image" + OfficeImageInfo.GetDefaultExtension(OfficeImageInfo.FromMimeType(contentType));
            }

            return new OfficeRemoteImage(current, bytes, fileName, contentType);
        }
    }

    private static HttpClient CreateClient() {
        var handler = new HttpClientHandler {
            AllowAutoRedirect = false,
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        };
        return new HttpClient(handler, disposeHandler: true) {
            Timeout = System.Threading.Timeout.InfiniteTimeSpan
        };
    }

#if NET8_0_OR_GREATER
    private static HttpClient CreatePinnedClient(
        bool allowPrivateNetworkAddresses) {
        var handler = new SocketsHttpHandler {
            AllowAutoRedirect = false,
            AutomaticDecompression = DecompressionMethods.GZip
                | DecompressionMethods.Deflate,
            PooledConnectionIdleTimeout = TimeSpan.FromMinutes(2),
            PooledConnectionLifetime = TimeSpan.FromMinutes(5),
            UseProxy = false
        };
        handler.ConnectCallback = async (context, cancellationToken) => {
            IPAddress[] addresses = await ResolveDestinationAsync(
                context.DnsEndPoint.Host, allowPrivateNetworkAddresses,
                cancellationToken).ConfigureAwait(false);
            Exception? lastFailure = null;
            foreach (IPAddress address in addresses) {
                var socket = new Socket(address.AddressFamily,
                    SocketType.Stream, ProtocolType.Tcp);
                try {
                    await socket.ConnectAsync(address,
                        context.DnsEndPoint.Port, cancellationToken)
                        .ConfigureAwait(false);
                    return new NetworkStream(socket, ownsSocket: true);
                } catch (Exception exception) when (exception
                           is SocketException
                           || exception is IOException) {
                    lastFailure = exception;
                    socket.Dispose();
                } catch {
                    socket.Dispose();
                    throw;
                }
            }

            throw new HttpRequestException(
                "No validated remote image address accepted the connection.",
                lastFailure);
        };
        return new HttpClient(handler, disposeHandler: true) {
            Timeout = System.Threading.Timeout.InfiniteTimeSpan
        };
    }
#endif

    private static async Task<byte[]> ReadBoundedAsync(Stream input, long maximumBytes, CancellationToken cancellationToken) {
        using var output = new MemoryStream();
        var buffer = new byte[BufferSize];
        long total = 0;
        while (true) {
            int read = await input.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;

            total += read;
            if (total > maximumBytes) {
                throw new InvalidDataException($"The remote image exceeds the {maximumBytes}-byte limit.");
            }

            await output.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
        }

        return output.ToArray();
    }

    private static string NormalizeImageContentType(string? contentType) {
        if (!OfficeImageInfo.TryNormalizeImageContentType(contentType, out string normalized)) {
            throw new InvalidDataException("The remote resource did not return an image content type.");
        }

        return normalized;
    }

    private static Uri ResolveRedirect(Uri current, Uri? location) {
        if (location == null) {
            throw new InvalidDataException("The remote image response contained a redirect without a location.");
        }

        Uri resolved = location.IsAbsoluteUri ? location : new Uri(current, location);
        ValidateHttpUri(resolved, nameof(location));
        if (!IsSameOrigin(current, resolved)) {
            throw new InvalidDataException("Remote image redirects must remain on the same origin.");
        }

        return resolved;
    }

    private static void ValidateHttpUri(Uri uri, string parameterName) {
        if (!uri.IsAbsoluteUri
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)) {
            throw new ArgumentException("Image URI must use HTTP or HTTPS.", parameterName);
        }
        if (!string.IsNullOrEmpty(uri.UserInfo)) {
            throw new ArgumentException("Image URI cannot contain embedded credentials.", parameterName);
        }
    }

    private static void ValidateDestinationPolicy(
        Uri uri,
        OfficeRemoteImageLoadOptions.Snapshot options) {
        string host = uri.IdnHost.TrimEnd('.');
        if (options.AllowedHosts.Count > 0 && !options.AllowedHosts.Contains(host)) {
            throw new InvalidDataException("The remote image host is outside the configured allowlist.");
        }
        if (!options.AllowPrivateNetworkAddresses
            && (uri.IsLoopback
            || string.Equals(host, "localhost", StringComparison.OrdinalIgnoreCase)
            || host.EndsWith(".localhost", StringComparison.OrdinalIgnoreCase))) {
            throw new InvalidDataException("Remote images cannot target localhost or non-public network addresses.");
        }
    }

    private static Task<IPAddress[]> ResolveDestinationAsync(
        Uri uri,
        bool allowPrivateNetworkAddresses,
        CancellationToken cancellationToken) => ResolveDestinationAsync(
            uri.IdnHost.TrimEnd('.'), allowPrivateNetworkAddresses,
            cancellationToken);

    private static async Task<IPAddress[]> ResolveDestinationAsync(
        string host,
        bool allowPrivateNetworkAddresses,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();

        IPAddress[] addresses;
        if (IPAddress.TryParse(host, out IPAddress? literal)) {
            addresses = new[] { literal };
        } else {
            Task<IPAddress[]> resolution = Dns.GetHostAddressesAsync(host);
            Task completed = await Task.WhenAny(
                resolution,
                Task.Delay(System.Threading.Timeout.Infinite, cancellationToken)).ConfigureAwait(false);
            if (completed != resolution) {
                cancellationToken.ThrowIfCancellationRequested();
            }
            addresses = await resolution.ConfigureAwait(false);
        }

        if (addresses.Length == 0
            || (!allowPrivateNetworkAddresses
                && addresses.Any(IsNonPublicAddress))) {
            throw new InvalidDataException("Remote images cannot target localhost or non-public network addresses.");
        }

        return addresses;
    }

    internal static HttpRequestMessage CreatePinnedRequest(Uri originalUri, IPAddress address) {
        if (originalUri == null) throw new ArgumentNullException(nameof(originalUri));
        if (address == null) throw new ArgumentNullException(nameof(address));

        var pinned = new UriBuilder(originalUri) { Host = address.ToString() }.Uri;
        var request = new HttpRequestMessage(HttpMethod.Get, pinned);
        string host = originalUri.IdnHost;
        if (host.IndexOf(':') >= 0 && host[0] != '[') {
            host = "[" + host + "]";
        }
        request.Headers.Host = originalUri.IsDefaultPort ? host : host + ":" + originalUri.Port;
        return request;
    }

    private static bool IsNonPublicAddress(IPAddress address) {
        if (address.IsIPv4MappedToIPv6) address = address.MapToIPv4();
        if (IPAddress.IsLoopback(address)
            || address.Equals(IPAddress.Any)
            || address.Equals(IPAddress.IPv6Any)
            || address.Equals(IPAddress.None)
            || address.Equals(IPAddress.IPv6None)) {
            return true;
        }

        byte[] bytes = address.GetAddressBytes();
        if (address.AddressFamily == AddressFamily.InterNetwork) {
            byte first = bytes[0];
            byte second = bytes[1];
            byte third = bytes[2];
            if (first == 0 || first == 10 || first == 127 || first >= 224) return true;
            if (first == 100 && second >= 64 && second <= 127) return true;
            if (first == 169 && second == 254) return true;
            if (first == 172 && second >= 16 && second <= 31) return true;
            if (first == 192 && second == 168) return true;
            if (first == 192 && second == 0 && third == 0) return true;
            if (first == 192 && second == 0 && third == 2) return true;
            if (first == 198 && (second == 18 || second == 19 || (second == 51 && third == 100))) return true;
            if (first == 203 && second == 0 && third == 113) return true;
            return false;
        }

        if (address.AddressFamily == AddressFamily.InterNetworkV6) {
            return address.IsIPv6LinkLocal
                || address.IsIPv6SiteLocal
                || address.IsIPv6Multicast
                || (bytes[0] & 0xFE) == 0xFC
                || (bytes[0] == 0x20 && bytes[1] == 0x01 && bytes[2] == 0x0D && bytes[3] == 0xB8);
        }
        return true;
    }

    private static bool IsSameOrigin(Uri left, Uri right) =>
        string.Equals(left.Scheme, right.Scheme, StringComparison.OrdinalIgnoreCase)
        && string.Equals(left.IdnHost, right.IdnHost, StringComparison.OrdinalIgnoreCase)
        && left.Port == right.Port;

    private static bool IsRedirect(HttpStatusCode statusCode) =>
        statusCode == HttpStatusCode.Moved
        || statusCode == HttpStatusCode.Redirect
        || statusCode == HttpStatusCode.SeeOther
        || statusCode == HttpStatusCode.TemporaryRedirect
        || (int)statusCode == 308;
}
