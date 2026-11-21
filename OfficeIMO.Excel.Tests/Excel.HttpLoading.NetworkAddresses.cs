using OfficeIMO.Excel;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("127.0.0.1")]
        [InlineData("localhost")]
        public async Task ExcelHttpLoadRefusesLoopbackLiteralAndResolvedNameUnlessAllowed(string host) {
            using var server = new LoopbackWorkbookServer(CreateRemoteWorkbookBytes());
            var uri = new Uri($"http://{host}:{server.Port}/workbook.xlsx");

            NotSupportedException refused = await Assert.ThrowsAsync<NotSupportedException>(() =>
                ExcelDocumentReader.OpenAsync(uri, httpOptions: new ExcelLoadOptions { SchemePolicy = ExcelUriSchemePolicy.HttpAndHttps }));
            Assert.Contains(nameof(ExcelLoadOptions.AllowPrivateNetworkAddresses), refused.Message, StringComparison.Ordinal);
            Assert.Equal(0, server.RequestCount);

            using var reader = await ExcelDocumentReader.OpenAsync(uri, httpOptions: new ExcelLoadOptions {
                SchemePolicy = ExcelUriSchemePolicy.HttpAndHttps,
                AllowPrivateNetworkAddresses = true
            });
            Assert.Equal(new[] { "Remote" }, reader.GetSheetNames());
            Assert.Equal(1, server.RequestCount);
        }

        [Fact]
        public async Task ExcelHttpLoadRefusesRedirectToPrivateAddressBeforeSendingIt() {
            using var server = new LoopbackWorkbookServer(CreateRemoteWorkbookBytes());
            int requests = 0;
            using var handler = new FakeWorkbookHttpMessageHandler((_, _) => {
                requests++;
                var redirect = new HttpResponseMessage(HttpStatusCode.Redirect);
                redirect.Headers.Location = new Uri($"http://127.0.0.1:{server.Port}/workbook.xlsx");
                return Task.FromResult(redirect);
            });

            await Assert.ThrowsAsync<NotSupportedException>(() =>
                ExcelDocumentReader.OpenAsync(
                    new Uri("http://example.test/workbook.xlsx"),
                    httpOptions: new ExcelLoadOptions {
                        SchemePolicy = ExcelUriSchemePolicy.HttpAndHttps,
                        HttpMessageHandler = handler
                    }));

            Assert.Equal(1, requests);
            Assert.Equal(0, server.RequestCount);
        }

        private sealed class LoopbackWorkbookServer : IDisposable {
            private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
            private readonly byte[] _workbook;
            private int _requestCount;

            internal LoopbackWorkbookServer(byte[] workbook) {
                _workbook = workbook;
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
                            "HTTP/1.1 200 OK\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n" +
                            $"Content-Length: {_workbook.Length}\r\nConnection: close\r\n\r\n");
                        await stream.WriteAsync(headers, 0, headers.Length).ConfigureAwait(false);
                        await stream.WriteAsync(_workbook, 0, _workbook.Length).ConfigureAwait(false);
                    }
                } catch (Exception exception) when (exception is SocketException or ObjectDisposedException or IOException) {
                    // The listener stops when the test disposes it.
                }
            }

            public void Dispose() => _listener.Stop();
        }
    }
}
