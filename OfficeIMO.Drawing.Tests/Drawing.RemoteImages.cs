using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRemoteImageTests {
    [Fact]
    public async Task LoadAsyncReturnsNormalizedImageDataWithDefensiveOutputs() {
        byte[] payload = { 0xFF, 0xD8, 0xFF, 0xD9 };
        using var server = new LoopbackImageServer(
            $"HTTP/1.1 200 OK\r\nContent-Type: image/jpg; charset=binary\r\nContent-Length: {payload.Length}\r\nConnection: close\r\n\r\n",
            payload);

        OfficeRemoteImage image = await OfficeRemoteImageLoader.LoadAsync(
            server.Url("logo.jpg"),
            new OfficeRemoteImageLoadOptions { AllowPrivateNetworkAddresses = true });

        Assert.Equal("image/jpeg", image.ContentType);
        Assert.Equal("logo.jpg", image.FileName);
        Assert.Equal(payload, image.ToBytes());
        byte[] returned = image.ToBytes();
        returned[0] = 0;
        Assert.Equal(0xFF, image.ToBytes()[0]);
        await server.Completion;
    }

    [Fact]
    public async Task LoadAsyncRejectsCrossOriginRedirectBeforeContactingTarget() {
        using var target = new LoopbackImageServer(
            "HTTP/1.1 200 OK\r\nContent-Type: image/png\r\nContent-Length: 1\r\nConnection: close\r\n\r\n",
            new byte[] { 1 });
        using var redirect = new LoopbackImageServer(
            $"HTTP/1.1 302 Found\r\nLocation: {target.Url("private.png")}\r\nConnection: close\r\n\r\n");

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            OfficeRemoteImageLoader.LoadAsync(
                redirect.Url("redirect.png"),
                new OfficeRemoteImageLoadOptions { AllowPrivateNetworkAddresses = true }));
        await redirect.Completion;
        await Task.Delay(100);
        Assert.Equal(0, target.RequestCount);
        target.Dispose();
        await target.Completion;
    }

    [Fact]
    public async Task LoadAsyncKeepsCookiesWithinOneLoad() {
        const string cookie = "officeimo_image_session=proof";
        string path = $"/{Guid.NewGuid():N}/";
        byte[] payload = { 0xFF, 0xD8, 0xFF, 0xD9 };
        string redirect = $"HTTP/1.1 302 Found\r\nLocation: {path}image.jpg\r\nSet-Cookie: {cookie}; Path={path}\r\nContent-Length: 0\r\nConnection: close\r\n\r\n";
        string success = $"HTTP/1.1 200 OK\r\nContent-Type: image/jpeg\r\nContent-Length: {payload.Length}\r\nConnection: close\r\n\r\n";
        using var server = new LoopbackImageServer((redirect, null), (success, payload), (success, payload));
        var options = new OfficeRemoteImageLoadOptions {
            AllowPrivateNetworkAddresses = true,
            Timeout = TimeSpan.FromSeconds(8)
        };
        string imageUrl = server.Url(path + "image.jpg");

        OfficeRemoteImage first = await OfficeRemoteImageLoader.LoadAsync(server.Url(path + "start.jpg"), options);
        OfficeRemoteImage second = await OfficeRemoteImageLoader.LoadAsync(imageUrl, options);
        await server.Completion;

        Assert.Equal(imageUrl, first.Source.AbsoluteUri);
        Assert.Equal(payload, first.ToBytes());
        Assert.Equal(payload, second.ToBytes());
        Assert.Equal(new string?[] { null, cookie, null }, server.RequestCookies);
    }

    [Fact]
    public async Task LoadAsyncEnforcesMaximumBytesWhileStreaming() {
        byte[] payload = Enumerable.Repeat((byte)0x41, 64).ToArray();
        using var server = new LoopbackImageServer(
            "HTTP/1.1 200 OK\r\nContent-Type: image/png\r\nConnection: close\r\n\r\n",
            payload);
        var options = new OfficeRemoteImageLoadOptions { MaximumBytes = 16, AllowPrivateNetworkAddresses = true };

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            OfficeRemoteImageLoader.LoadAsync(server.Url("large.png"), options));
        await server.Completion;
    }

    [Fact]
    public async Task LoadAsyncHonorsCallerCancellation() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            OfficeRemoteImageLoader.LoadAsync("http://127.0.0.1:1/image.png", cancellationToken: cancellation.Token));
    }

    [Theory]
    [InlineData("127.0.0.1")]
    [InlineData("[::ffff:127.0.0.1]")]
    [InlineData("localhost")]
    public async Task LoadAsyncRefusesLoopbackLiteralAndResolvedNameByDefault(string host) {
        using var server = new LoopbackImageServer(
            "HTTP/1.1 200 OK\r\nContent-Type: image/png\r\nContent-Length: 1\r\nConnection: close\r\n\r\n",
            new byte[] { 1 });

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            OfficeRemoteImageLoader.LoadAsync($"http://{host}:{server.Port}/private.png"));
        Assert.Equal(0, server.RequestCount);
    }

    [Theory]
    [InlineData("127.0.0.1", false)]
    [InlineData("10.1.2.3", false)]
    [InlineData("100.64.0.1", false)]
    [InlineData("169.254.169.254", false)]
    [InlineData("172.16.0.1", false)]
    [InlineData("192.168.1.1", false)]
    [InlineData("0.0.0.0", false)]
    [InlineData("::", false)]
    [InlineData("::1", false)]
    [InlineData("fe80::1", false)]
    [InlineData("fd00::1", false)]
    [InlineData("::ffff:169.254.169.254", false)]
    [InlineData("::ffff:0:10.0.0.1", false)]
    [InlineData("64:ff9b::a00:1", false)]
    [InlineData("2002:c0a8:101::1", false)]
    [InlineData("93.184.215.14", true)]
    [InlineData("2606:4700::1111", true)]
    [InlineData("64:ff9b::808:808", true)]
    public void NetworkAddressPolicyClassifiesDestinations(string address, bool expectedPublic) {
        Assert.Equal(expectedPublic, OfficeNetworkAddressPolicy.IsPublic(IPAddress.Parse(address)));
    }

    [Fact]
    public void LegacyPinnedRequestUsesCheckedAddressAndPreservesHostHeader() {
        using HttpRequestMessage request = OfficeNetworkAddressPolicy.CreatePinnedRequest(
            HttpMethod.Get,
            new Uri("http://images.example.test:8080/assets/logo.png?size=2"),
            IPAddress.Parse("203.0.113.8"));

        Assert.Equal("203.0.113.8", request.RequestUri!.Host);
        Assert.Equal("/assets/logo.png?size=2", request.RequestUri.PathAndQuery);
        Assert.Equal("images.example.test:8080", request.Headers.Host);
    }

    private sealed class LoopbackImageServer : IDisposable {
        private readonly TcpListener _listener;
        private readonly List<string?> _requestCookies = new();
        private bool _disposed;

        internal LoopbackImageServer(string headers, byte[]? payload = null) : this((headers, payload)) {
        }

        internal LoopbackImageServer(params (string Headers, byte[]? Payload)[] responses) {
            _listener = new TcpListener(IPAddress.Loopback, 0);
            _listener.Start();
            Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
            Completion = ServeAsync(responses);
        }

        internal int Port { get; }
        internal int RequestCount { get; private set; }
        internal IReadOnlyList<string?> RequestCookies => _requestCookies;
        internal Task Completion { get; }

        internal string Url(string path) => $"http://127.0.0.1:{Port}/{path.TrimStart('/')}";

        private async Task ServeAsync((string Headers, byte[]? Payload)[] responses) {
            try {
                foreach ((string headers, byte[]? payload) in responses) {
                    using TcpClient client = await _listener.AcceptTcpClientAsync();
                    RequestCount++;
                    using NetworkStream stream = client.GetStream();
                    using (var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, leaveOpen: true)) {
                        string? cookie = null;
                        string? line;
                        while (!string.IsNullOrEmpty(line = await reader.ReadLineAsync())) {
                            if (line.StartsWith("Cookie:", StringComparison.OrdinalIgnoreCase)) {
                                cookie = line.Substring(7).Trim();
                            }
                        }
                        _requestCookies.Add(cookie);
                    }

                    byte[] headerBytes = Encoding.ASCII.GetBytes(headers);
                    await stream.WriteAsync(headerBytes, 0, headerBytes.Length);
                    if (payload is { Length: > 0 }) {
                        await stream.WriteAsync(payload, 0, payload.Length);
                    }
                    await stream.FlushAsync();
                }
            } catch (SocketException) when (_disposed) {
                // Normal disposal path for a listener that did not receive a request.
            } catch (ObjectDisposedException) when (_disposed) {
                // Normal disposal path for a listener that did not receive a request.
            }
        }

        public void Dispose() {
            if (_disposed) return;
            _disposed = true;
            _listener.Stop();
        }
    }
}
