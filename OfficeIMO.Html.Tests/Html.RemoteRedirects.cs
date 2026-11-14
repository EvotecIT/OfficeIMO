using System.Net;
using System.Net.Sockets;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    private const string RedirectTestPngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII=";

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    public async Task HtmlToWord_BuiltInHttpClientChecksRedirectHostBeforeFetching(bool stylesheet, bool allowed, bool useUrlTransform) {
        byte[] payload = stylesheet
            ? Encoding.UTF8.GetBytes("p { color:#123456; }")
            : Convert.FromBase64String(RedirectTestPngBase64);
        string contentType = stylesheet ? "text/css" : "image/png";
        string extension = stylesheet ? "css" : "png";
        using var destination = new RedirectTestServer();
        Task destinationResponse = destination.SendAsync($"200 OK\r\nContent-Type: {contentType}", payload);
        using var origin = new RedirectTestServer();
        string destinationHost = allowed && !useUrlTransform ? "localhost" : "127.0.0.1";
        Task originResponse = origin.SendAsync($"302 Found\r\nLocation: //{destinationHost}:{destination.Port}/resource.{extension}");
        string source = $"http://localhost:{origin.Port}/start.{extension}";

        var options = new HtmlWordExportOptions {
            AllowDocumentStylesheetLinks = true,
            ImageProcessing = WordHtmlImageProcessingMode.Embed,
            ResourceTimeout = TimeSpan.FromSeconds(8)
        };
        if (useUrlTransform) {
            options.ResourceUrlPolicy.ResolvedUrlTransform = url => {
                var uri = new Uri(url);
                if (uri.Host == "localhost") return url;
                return allowed ? new UriBuilder(uri) { Host = "localhost" }.Uri.AbsoluteUri : null;
            };
        } else {
            options.AllowedImageHosts.Add("localhost");
        }
        options.AllowedStylesheetHosts.Add("localhost");
        string html = stylesheet
            ? $"<link rel='stylesheet' href='{source}'><p>Body text</p>"
            : $"<p><img src='{source}' alt='Image fallback'></p>";

        HtmlToWordResult result = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ConvertToWordWithEvidenceAsync(options);
        using WordDocument document = result.Value;
        origin.Dispose();
        await originResponse;
        destination.Dispose();
        await destinationResponse;

        Assert.True(origin.WasRequested);
        Assert.Equal(allowed, destination.WasRequested);
        if (!allowed) {
            string code = stylesheet ? "stylesheet-resource-rejected-by-policy" : "image-resource-rejected-by-policy";
            Assert.Contains(result.Report.NativeDiagnostics, diagnostic => diagnostic.Code == code);
        }
        if (allowed) {
            Assert.Equal($"localhost:{destination.Port}", destination.RequestHost);
        }
        if (stylesheet) {
            Assert.Equal(allowed, document.Paragraphs.First().GetRuns().First().Color == OfficeColor.Parse("123456"));
        } else if (allowed) {
            Assert.Single(document.Images);
        } else {
            Assert.Empty(document.Images);
            Assert.Contains(document.Paragraphs, paragraph => paragraph.Text.Contains("Image fallback"));
        }
    }

    [Theory]
    [InlineData(false)]
#if NETFRAMEWORK
    [InlineData(true)]
#endif
    public async Task HtmlToWord_BuiltInHttpClientKeepsCookiesWithinOneConversion(bool useLegacyCookieHeader) {
        const string host = "127.0.0.1";
        const string cookie = "officeimo_session=proof";
        string path = $"/{Guid.NewGuid():N}/";
        string setCookie = useLegacyCookieHeader
            ? $"Set-Cookie2: {cookie}; Version=1; Path={path}"
            : $"Set-Cookie: {cookie}; Path={path}";
        string expectedCookie = useLegacyCookieHeader ? $"$Version=1; {cookie}; $Path={path}" : cookie;
        byte[] png = Convert.FromBase64String(RedirectTestPngBase64);
        using var origin = new RedirectTestServer();
        using var stylesheet = new RedirectTestServer();
        using var image = new RedirectTestServer();
        string startUrl = $"http://{host}:{origin.Port}{path}start.css";
        string stylesheetUrl = $"http://{host}:{stylesheet.Port}{path}style.css";
        string imageUrl = $"http://{host}:{image.Port}{path}image.png";
        Task originResponse = origin.SendAsync($"302 Found\r\nLocation: {stylesheetUrl}\r\n{setCookie}");
        Task stylesheetResponse = stylesheet.SendAsync("200 OK\r\nContent-Type: text/css", Encoding.UTF8.GetBytes("p { color:#123456; }"));
        Task firstImageResponse = image.SendAsync("200 OK\r\nContent-Type: image/png", png);

        var options = new HtmlWordExportOptions {
            AllowDocumentStylesheetLinks = true,
            ImageProcessing = WordHtmlImageProcessingMode.Embed,
            ResourceTimeout = TimeSpan.FromSeconds(8)
        };
        options.AllowedImageHosts.Add(host);
        options.AllowedStylesheetHosts.Add(host);
        string imageHtml = $"<p><img src='{imageUrl}' alt='Image fallback'></p>";
        using WordDocument first = (await OfficeIMO.Html.HtmlConversionDocument
            .Parse($"<link rel='stylesheet' href='{startUrl}'><p>Styled text</p>{imageHtml}")
            .ToWordAsync(options)).Value;
        origin.Dispose();
        stylesheet.Dispose();
        Assert.Single(first.Images);
        await Task.WhenAll(originResponse, stylesheetResponse, firstImageResponse);
        Assert.Null(origin.RequestCookie);
        Assert.Equal(expectedCookie, stylesheet.RequestCookie);
        Assert.Equal(expectedCookie, image.RequestCookie);

        Task secondImageResponse = image.SendAsync("200 OK\r\nContent-Type: image/png", png);
        using WordDocument second = (await OfficeIMO.Html.HtmlConversionDocument.Parse(imageHtml)
            .ToWordAsync(options)).Value;
        image.Dispose();
        await secondImageResponse;
        Assert.Single(second.Images);
        Assert.Null(image.RequestCookie);
    }

    private sealed class RedirectTestServer : IDisposable {
        // Clients resolve "localhost" to ::1 before 127.0.0.1. Listening on both loopbacks keeps
        // a refused IPv6 attempt (about two seconds on Windows) out of each request's timeout.
        private readonly TcpListener _listener;
        private readonly TcpListener? _ipv6Listener;
        private Task<TcpClient>? _pendingAccept;
        private Task<TcpClient>? _pendingIpv6Accept;
        private volatile bool _disposed;

        internal RedirectTestServer() {
            for (int attempt = 0; ; attempt++) {
                _listener = new TcpListener(IPAddress.Loopback, 0);
                _listener.Start();
                Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
                if (!Socket.OSSupportsIPv6) return;
                try {
                    _ipv6Listener = new TcpListener(IPAddress.IPv6Loopback, Port);
                    _ipv6Listener.Start();
                    return;
                } catch (SocketException) when (attempt < 20) {
                    _ipv6Listener = null;
                    _listener.Stop();
                }
            }
        }

        internal int Port { get; }
        internal bool WasRequested { get; private set; }
        internal string? RequestHost { get; private set; }
        internal string? RequestCookie { get; private set; }

        internal async Task SendAsync(string statusAndHeaders, byte[]? payload = null) {
            try {
                using TcpClient client = await AcceptAsync();
                WasRequested = true;
                RequestCookie = null;
                using NetworkStream stream = client.GetStream();
                using var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, leaveOpen: true);
                string? line;
                while (!string.IsNullOrEmpty(line = await reader.ReadLineAsync())) {
                    if (line.StartsWith("Host:", StringComparison.OrdinalIgnoreCase)) {
                        RequestHost = line.Substring(5).Trim();
                    } else if (line.StartsWith("Cookie:", StringComparison.OrdinalIgnoreCase)) {
                        RequestCookie = line.Substring(7).Trim();
                    }
                }

                byte[] header = Encoding.ASCII.GetBytes(
                    $"HTTP/1.1 {statusAndHeaders}\r\nContent-Length: {payload?.Length ?? 0}\r\nConnection: close\r\n\r\n");
                await stream.WriteAsync(header, 0, header.Length);
                if (payload != null) {
                    await stream.WriteAsync(payload, 0, payload.Length);
                }
            } catch (SocketException) when (_disposed) {
            } catch (ObjectDisposedException) when (_disposed) {
            }
        }

        private async Task<TcpClient> AcceptAsync() {
            _pendingAccept ??= _listener.AcceptTcpClientAsync();
            if (_ipv6Listener == null) {
                Task<TcpClient> accepted = _pendingAccept;
                _pendingAccept = null;
                return await accepted;
            }

            // The losing accept stays pending for the next request on this server.
            _pendingIpv6Accept ??= _ipv6Listener.AcceptTcpClientAsync();
            Task<TcpClient> first = await Task.WhenAny(_pendingAccept, _pendingIpv6Accept);
            if (first == _pendingAccept) {
                _pendingAccept = null;
            } else {
                _pendingIpv6Accept = null;
            }
            return await first;
        }

        public void Dispose() {
            _disposed = true;
            _listener.Stop();
            _ipv6Listener?.Stop();
            Observe(_pendingAccept);
            Observe(_pendingIpv6Accept);
        }

        private static void Observe(Task? pending) {
            pending?.ContinueWith(static task => _ = task.Exception, TaskContinuationOptions.OnlyOnFaulted);
        }
    }
}
