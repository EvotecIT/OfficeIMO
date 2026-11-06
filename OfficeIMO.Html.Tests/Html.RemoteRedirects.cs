using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
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
            : Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII=");
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

    private sealed class RedirectTestServer : IDisposable {
        private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
        private volatile bool _disposed;

        internal RedirectTestServer() {
            _listener.Start();
            Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
        }

        internal int Port { get; }
        internal bool WasRequested { get; private set; }
        internal string? RequestHost { get; private set; }

        internal async Task SendAsync(string statusAndHeaders, byte[]? payload = null) {
            try {
                using TcpClient client = await _listener.AcceptTcpClientAsync();
                WasRequested = true;
                using NetworkStream stream = client.GetStream();
                using var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, leaveOpen: true);
                string? line;
                while (!string.IsNullOrEmpty(line = await reader.ReadLineAsync())) {
                    if (line.StartsWith("Host:", StringComparison.OrdinalIgnoreCase)) {
                        RequestHost = line.Substring(5).Trim();
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

        public void Dispose() {
            _disposed = true;
            _listener.Stop();
        }
    }
}
