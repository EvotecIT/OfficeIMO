using System.Net;
using System.Net.Http;
#if NET8_0_OR_GREATER
using System.Net.Sockets;
#endif

namespace OfficeIMO.Workflows;

/// <summary>The local CUPS scheduler, reached over IPP through its domain socket or the loopback port.</summary>
/// <remarks>
/// A <c>CUPS_SERVER</c> that names another host leaves the scheduler unavailable here, so the CUPS command-line tools,
/// which honor that setting, deliver the job instead.
/// </remarks>
internal sealed class CupsScheduler {
    private static readonly string[] SocketPaths = { "/private/var/run/cupsd", "/var/run/cupsd", "/run/cups/cups.sock", "/var/run/cups/cups.sock" };
    private readonly Uri? _baseUri;
    private readonly Func<HttpMessageHandler>? _createHandler;
    private readonly Func<string, bool, Uri>? _queueUri;

    internal CupsScheduler(Uri? baseUri, Func<HttpMessageHandler>? createHandler, Func<string, bool, Uri>? queueUri = null) {
        _baseUri = baseUri;
        _createHandler = createHandler;
        _queueUri = queueUri;
    }

    internal static CupsScheduler Local { get; } = FromEnvironment(Environment.GetEnvironmentVariable("CUPS_SERVER"), File.Exists);

    internal bool IsAvailable => _baseUri is not null;

    /// <summary>A client addressed to the scheduler itself, for CUPS-Get-Printers and CUPS-Get-Default.</summary>
    internal IppPrintClient CreateSchedulerClient() => IppPrinterEndpoint.ForUri("CUPS", Require(), _createHandler).CreateClient();

    /// <summary>A queue on the scheduler; classes live under /classes/ and printers under /printers/.</summary>
    internal IppPrinterEndpoint Queue(string queue, bool isClass) {
        Uri uri = _queueUri?.Invoke(queue, isClass) ?? new Uri(Require(), (isClass ? "classes/" : "printers/") + Uri.EscapeDataString(queue));
        return IppPrinterEndpoint.ForCupsQueue(queue, uri, _createHandler);
    }

    private Uri Require() => _baseUri ?? throw new PlatformNotSupportedException("The CUPS scheduler is not reachable over IPP on this host.");

    internal static CupsScheduler FromEnvironment(string? server, Func<string, bool> exists) {
        if (server is { Length: > 0 }) {
            if (server.StartsWith("/", StringComparison.Ordinal)) return Socket(server);
            string host = server, port = "631";
            int colon = server.LastIndexOf(':');
            if (colon > 0 && server.IndexOf(']') < colon) { host = server.Substring(0, colon); port = server.Substring(colon + 1); }
            host = host.Trim('[', ']');
            bool loopback = host.Equals("localhost", StringComparison.OrdinalIgnoreCase) || (IPAddress.TryParse(host, out IPAddress? address) && IPAddress.IsLoopback(address));
            return loopback && int.TryParse(port, out int number) && number is > 0 and < 65536
                ? new CupsScheduler(new Uri("ipp://localhost:" + number + "/"), null)
                : new CupsScheduler(null, null);
        }
        foreach (string path in SocketPaths) if (exists(path)) return Socket(path);
        return new CupsScheduler(new Uri("ipp://localhost:631/"), null);
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
