using System.Collections.Concurrent;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace OfficeIMO.Tests;

internal sealed class RuntimeHttpFixture : IAsyncDisposable {
    private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
    private readonly CancellationTokenSource _stop = new();
    private readonly Func<string, CancellationToken, Task<Reply>> _respond;
    private readonly List<Task> _clients = new();
    private readonly Task _accept;
    internal ConcurrentQueue<string> Requests { get; } = new();
    internal Uri Origin { get; }

    internal RuntimeHttpFixture(Func<string, CancellationToken, Task<Reply>> respond) {
        _respond = respond;
        _listener.Start();
        Origin = new Uri("http://127.0.0.1:" + ((IPEndPoint)_listener.LocalEndpoint).Port + "/");
        _accept = AcceptAsync();
    }

    private async Task AcceptAsync() {
        try {
            while (!_stop.IsCancellationRequested) {
                var client = await _listener.AcceptTcpClientAsync(_stop.Token);
                _clients.Add(ServeAsync(client));
            }
        } catch (OperationCanceledException) when (_stop.IsCancellationRequested) { }
    }

    private async Task ServeAsync(TcpClient client) {
        using (client) {
            try {
                var stream = client.GetStream();
                using var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, leaveOpen: true);
                string line = await reader.ReadLineAsync(_stop.Token) ?? throw new IOException("No request.");
                string path = line.Split(' ')[1];
                Requests.Enqueue(path);
                while (!string.IsNullOrEmpty(await reader.ReadLineAsync(_stop.Token))) { }
                Reply reply = await _respond(path, _stop.Token);
                string header = $"HTTP/1.1 {reply.Status} Test\r\nConnection: close\r\nContent-Type: {reply.ContentType}\r\n" +
                    (reply.Chunked ? "Transfer-Encoding: chunked\r\n" : $"Content-Length: {reply.Content.Length}\r\n") + reply.Headers + "\r\n";
                await stream.WriteAsync(Encoding.ASCII.GetBytes(header), _stop.Token);
                if (reply.Chunked) await stream.WriteAsync(Encoding.ASCII.GetBytes(reply.Content.Length.ToString("X") + "\r\n"), _stop.Token);
                await stream.WriteAsync(reply.Content, _stop.Token);
                if (reply.Chunked) await stream.WriteAsync(Encoding.ASCII.GetBytes("\r\n0\r\n\r\n"), _stop.Token);
            } catch (OperationCanceledException) when (_stop.IsCancellationRequested) { }
            catch (IOException) { /* Expected when the worker cancels or rejects a response. */ }
        }
    }

    public async ValueTask DisposeAsync() {
        _stop.Cancel();
        _listener.Stop();
        await _accept;
        await Task.WhenAll(_clients);
        _stop.Dispose();
    }

    internal sealed record Reply(byte[] Content, string ContentType = "text/javascript", int Status = 200, string Headers = "", bool Chunked = false) {
        internal static Reply Text(string content, string contentType = "text/javascript", int status = 200, string headers = "", bool chunked = false) => new(Encoding.UTF8.GetBytes(content), contentType, status, headers, chunked);
    }
}
