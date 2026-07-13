using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRemoteImageTests {
    [Fact]
    public async Task LoadAsyncReturnsNormalizedImageDataWithDefensiveOutputs() {
        byte[] payload = { 0xFF, 0xD8, 0xFF, 0xD9 };
        using var server = new SingleResponseServer(
            $"HTTP/1.1 200 OK\r\nContent-Type: image/jpg; charset=binary\r\nContent-Length: {payload.Length}\r\nConnection: close\r\n\r\n",
            payload);

        OfficeRemoteImage image = await OfficeRemoteImageLoader.LoadAsync(server.Url("logo.jpg"));

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
        using var target = new SingleResponseServer(
            "HTTP/1.1 200 OK\r\nContent-Type: image/png\r\nContent-Length: 1\r\nConnection: close\r\n\r\n",
            new byte[] { 1 });
        using var redirect = new SingleResponseServer(
            $"HTTP/1.1 302 Found\r\nLocation: {target.Url("private.png")}\r\nConnection: close\r\n\r\n");

        await Assert.ThrowsAsync<InvalidDataException>(() =>
            OfficeRemoteImageLoader.LoadAsync(redirect.Url("redirect.png")));
        await redirect.Completion;
        await Task.Delay(100);
        Assert.Equal(0, target.RequestCount);
        target.Dispose();
        await target.Completion;
    }

    [Fact]
    public async Task LoadAsyncEnforcesMaximumBytesWhileStreaming() {
        byte[] payload = Enumerable.Repeat((byte)0x41, 64).ToArray();
        using var server = new SingleResponseServer(
            "HTTP/1.1 200 OK\r\nContent-Type: image/png\r\nConnection: close\r\n\r\n",
            payload);
        var options = new OfficeRemoteImageLoadOptions { MaximumBytes = 16 };

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

    private sealed class SingleResponseServer : IDisposable {
        private readonly TcpListener _listener;
        private bool _disposed;

        internal SingleResponseServer(string headers, byte[]? payload = null) {
            _listener = new TcpListener(IPAddress.Loopback, 0);
            _listener.Start();
            Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
            Completion = ServeAsync(headers, payload ?? Array.Empty<byte>());
        }

        internal int Port { get; }
        internal int RequestCount { get; private set; }
        internal Task Completion { get; }

        internal string Url(string path) => $"http://127.0.0.1:{Port}/{path.TrimStart('/')}";

        private async Task ServeAsync(string headers, byte[] payload) {
            try {
                using TcpClient client = await _listener.AcceptTcpClientAsync();
                RequestCount++;
                using NetworkStream stream = client.GetStream();
                using (var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, leaveOpen: true)) {
                    string? line;
                    while (!string.IsNullOrEmpty(line = await reader.ReadLineAsync())) { }
                }

                byte[] headerBytes = Encoding.ASCII.GetBytes(headers);
                await stream.WriteAsync(headerBytes, 0, headerBytes.Length);
                if (payload.Length > 0) {
                    await stream.WriteAsync(payload, 0, payload.Length);
                }
                await stream.FlushAsync();
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
