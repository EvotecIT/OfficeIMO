using OfficeIMO.Reader;
using OfficeIMO.Reader.Web;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderWebNetworkAddressTests {
    [Theory]
    [InlineData("127.0.0.1")]
    [InlineData("localhost")]
    public async Task WebReaderWithoutClientRefusesLoopbackLiteralAndResolvedNameUnlessAllowed(string host) {
        using var server = new LoopbackTextServer("loopback fixture");
        var uri = new Uri($"http://{host}:{server.Port}/fixture.txt");

        OfficeDocumentWebReader defaultReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader();
        await Assert.ThrowsAsync<ReaderWebPolicyException>(() => defaultReader.ReadDocumentAsync(uri));
        Assert.Equal(0, server.RequestCount);

        OfficeDocumentWebReader trustedReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(
            new ReaderWebOptions { AllowPrivateNetworkAddresses = true });
        OfficeDocumentReadResult result = await trustedReader.ReadDocumentAsync(uri);
        Assert.Contains("loopback fixture", result.Markdown, StringComparison.Ordinal);
        Assert.Equal(1, server.RequestCount);
    }

    [Fact]
    public async Task WebReaderRefusesRedirectToPrivateAddressBeforeSendingIt() {
        using var server = new LoopbackTextServer("not reached");
        int calls = 0;
        using var httpClient = new HttpClient(new RedirectingHandler(() => {
            calls++;
            var redirect = new HttpResponseMessage(HttpStatusCode.Found);
            redirect.Headers.Location = new Uri($"http://127.0.0.1:{server.Port}/fixture.txt");
            return redirect;
        }));
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        ReaderWebPolicyException refused = await Assert.ThrowsAsync<ReaderWebPolicyException>(() =>
            webReader.ReadDocumentAsync(new Uri("http://example.test/start.txt")));

        Assert.Equal("127.0.0.1", refused.TargetUri.Host);
        Assert.Equal(1, calls);
        Assert.Equal(0, server.RequestCount);
    }

    [Fact]
    public async Task WebReaderRefusesCrossOriginRedirectWhenTheClientSendsDefaultCredentials() {
        var requestedHosts = new List<string>();
        using var httpClient = new HttpClient(new RecordingHandler(request => {
            requestedHosts.Add(request.RequestUri!.Host);
            if (request.RequestUri.Host == "docs.example.test") {
                var redirect = new HttpResponseMessage(HttpStatusCode.Found);
                redirect.Headers.Location = new Uri("https://other.example.test/file.txt");
                return redirect;
            }
            return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("reached") };
        }));
        httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", "secret");
        OfficeDocumentWebReader webReader = OfficeIMO.Reader.Tests.ReaderTestReaders.All.CreateWebReader(httpClient);

        ReaderWebPolicyException refused = await Assert.ThrowsAsync<ReaderWebPolicyException>(() =>
            webReader.ReadDocumentAsync(new Uri("https://docs.example.test/start.txt")));

        Assert.Equal("other.example.test", refused.TargetUri.Host);
        Assert.Equal(new[] { "docs.example.test" }, requestedHosts);
    }

    private sealed class RecordingHandler : HttpMessageHandler {
        private readonly Func<HttpRequestMessage, HttpResponseMessage> _respond;

        internal RecordingHandler(Func<HttpRequestMessage, HttpResponseMessage> respond) => _respond = respond;

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            Task.FromResult(_respond(request));
    }

    private sealed class RedirectingHandler : HttpMessageHandler {
        private readonly Func<HttpResponseMessage> _respond;

        internal RedirectingHandler(Func<HttpResponseMessage> respond) => _respond = respond;

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            Task.FromResult(_respond());
    }

    private sealed class LoopbackTextServer : IDisposable {
        private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
        private readonly byte[] _body;
        private int _requestCount;

        internal LoopbackTextServer(string body) {
            _body = Encoding.UTF8.GetBytes(body);
            _listener.Start();
            Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
            _ = ServeAsync();
        }

        internal int Port { get; }
        internal int RequestCount => Volatile.Read(ref _requestCount);

        private async Task ServeAsync() {
            try {
                while (true) {
                    using TcpClient client = await _listener.AcceptTcpClientAsync().ConfigureAwait(false);
                    Interlocked.Increment(ref _requestCount);
                    using NetworkStream stream = client.GetStream();
                    using (var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, leaveOpen: true)) {
                        while (!string.IsNullOrEmpty(await reader.ReadLineAsync().ConfigureAwait(false))) {
                        }
                    }
                    byte[] headers = Encoding.ASCII.GetBytes(
                        $"HTTP/1.1 200 OK\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Length: {_body.Length}\r\nConnection: close\r\n\r\n");
                    await stream.WriteAsync(headers, 0, headers.Length).ConfigureAwait(false);
                    await stream.WriteAsync(_body, 0, _body.Length).ConfigureAwait(false);
                }
            } catch (Exception exception) when (exception is SocketException or ObjectDisposedException or IOException) {
                // The listener stops when the test disposes it.
            }
        }

        public void Dispose() => _listener.Stop();
    }
}
