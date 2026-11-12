using System.Net;
using System.Net.Http;
#if NET8_0_OR_GREATER
using System.Net.Sockets;
#endif

namespace OfficeIMO.Workflows;

/// <summary>The CUPS scheduler, reached over IPP through its domain socket, the loopback port, or the host <c>CUPS_SERVER</c> names.</summary>
/// <remarks>
/// A <c>CUPS_SERVER</c> that names another host leaves queue listing and submission to the CUPS command-line tools, which honor that
/// setting; job status and cancellation for their receipts go to that scheduler over IPP.
/// </remarks>
internal sealed class CupsScheduler {
    private static readonly string[] SocketPaths = { "/private/var/run/cupsd", "/var/run/cupsd", "/run/cups/cups.sock", "/var/run/cups/cups.sock" };
    private readonly Uri? _baseUri;
    private readonly Uri? _jobsUri;
    private readonly Func<HttpMessageHandler>? _createHandler;
    private readonly Func<string, bool, Uri>? _queueUri;

    internal CupsScheduler(Uri? baseUri, Func<HttpMessageHandler>? createHandler, Func<string, bool, Uri>? queueUri = null, Uri? jobsUri = null) {
        _baseUri = baseUri;
        _createHandler = createHandler;
        _queueUri = queueUri;
        _jobsUri = jobsUri ?? baseUri;
    }

    internal static CupsScheduler Local { get; } = FromEnvironment(Environment.GetEnvironmentVariable("CUPS_SERVER"), File.Exists);

    /// <summary>True when queues are listed and jobs submitted over IPP; otherwise the command-line tools do both.</summary>
    internal bool IsAvailable => _baseUri is not null;

    /// <summary>A client addressed to the scheduler itself, for CUPS-Get-Printers and CUPS-Get-Default.</summary>
    internal IppPrintClient CreateSchedulerClient() => IppPrinterEndpoint.ForUri("CUPS", Require(), _createHandler).CreateClient();

    /// <summary>A queue on the scheduler; classes live under /classes/ and printers under /printers/.</summary>
    internal IppPrinterEndpoint Queue(string queue, bool isClass) => QueueAt(Require(), queue, isClass);

    /// <summary>The queue on the scheduler that owns its jobs, including jobs <c>lp</c> submitted to a remote <c>CUPS_SERVER</c>.</summary>
    /// <remarks>The scheduler finds a job by its number; the queue path only scopes the request.</remarks>
    internal IppPrinterEndpoint JobQueue(string queue) => QueueAt(
        _jobsUri ?? throw new PlatformNotSupportedException("Job tracking needs the CUPS scheduler over IPP, which is not reachable on this host."), queue, isClass: false);

    private IppPrinterEndpoint QueueAt(Uri scheduler, string queue, bool isClass) {
        Uri uri = _queueUri?.Invoke(queue, isClass) ?? new Uri(scheduler, (isClass ? "classes/" : "printers/") + Uri.EscapeDataString(queue));
        return IppPrinterEndpoint.ForCupsQueue(queue, uri, _createHandler);
    }

    private Uri Require() => _baseUri ?? throw new PlatformNotSupportedException("The CUPS scheduler is not reachable over IPP on this host.");

    internal static CupsScheduler FromEnvironment(string? server, Func<string, bool> exists) {
        if (server is { Length: > 0 }) {
            if (server.StartsWith("/", StringComparison.Ordinal)) return Socket(server);
            if (ServerUri(server) is not { } uri) return new CupsScheduler(null, null);
            return uri.IsLoopback ? new CupsScheduler(uri, null) : new CupsScheduler(null, null, jobsUri: uri);
        }
        foreach (string path in SocketPaths) if (exists(path)) return Socket(path);
        return new CupsScheduler(new Uri("ipp://localhost:631/"), null);
    }

    /// <summary>
    /// The scheduler URI for a <c>CUPS_SERVER</c> value of the form <c>host</c>, <c>host:port</c>, <c>[ipv6]</c> or <c>[ipv6]:port</c>,
    /// optionally followed by <c>/version=1.1</c>; null when the value is not a usable address.
    /// </summary>
    internal static Uri? ServerUri(string server) {
        int slash = server.IndexOf('/');
        if (slash >= 0) server = server.Substring(0, slash);
        string host = server, port = "631";
        if (server.StartsWith("[", StringComparison.Ordinal)) {
            int close = server.IndexOf(']');
            if (close < 0) return null;
            host = server.Substring(1, close - 1);
            string rest = server.Substring(close + 1);
            if (rest.Length > 0) {
                if (rest[0] != ':') return null;
                port = rest.Substring(1);
            }
            if (!IPAddress.TryParse(host, out IPAddress? address) || address.AddressFamily != System.Net.Sockets.AddressFamily.InterNetworkV6) return null;
        } else if (server.IndexOf(':') is int colon and >= 0) {
            // An unbracketed IPv6 address has more than one colon and no port.
            if (server.LastIndexOf(':') != colon) {
                if (!IPAddress.TryParse(server, out _)) return null;
            } else {
                host = server.Substring(0, colon);
                port = server.Substring(colon + 1);
            }
        }
        if (!int.TryParse(port, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int number) || number is < 1 or > 65535) return null;
        if (Uri.CheckHostName(host) is UriHostNameType.Unknown or UriHostNameType.Basic) return null;
        return new UriBuilder("ipp", host, number, "/").Uri;
    }

    private static CupsScheduler Socket(string path) {
#if NET8_0_OR_GREATER
        return new CupsScheduler(new Uri("ipp://localhost/"), () => new SocketsHttpHandler {
            AllowAutoRedirect = false, UseCookies = false, UseProxy = false, AutomaticDecompression = DecompressionMethods.None,
            ConnectCallback = async (_, token) => {
                var socket = new System.Net.Sockets.Socket(AddressFamily.Unix, SocketType.Stream, ProtocolType.Unspecified);
                try {
                    await socket.ConnectAsync(new UnixDomainSocketEndPoint(path), token).ConfigureAwait(false);
                    return new NetworkStream(socket, ownsSocket: true);
                } catch {
                    socket.Dispose();
                    throw;
                }
            }
        });
#else
        // Older frameworks have no domain-socket HTTP connections; the scheduler's loopback port may still be open.
        _ = path;
        return new CupsScheduler(new Uri("ipp://localhost:631/"), null);
#endif
    }
}
