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
    internal ConcurrentQueue<ReceivedRequest> Received { get; } = new();
    internal Func<ReceivedRequest, CancellationToken, Task<Reply>>? RespondToRequest { get; set; }
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
                using var input = new MemoryStream();
                var one = new byte[1];
                while (input.Length < 32768) {
                    if (await stream.ReadAsync(one, _stop.Token) != 1) throw new IOException("No request.");
                    input.WriteByte(one[0]);
                    var bytes = input.GetBuffer();
                    int n = (int)input.Length;
                    if (n >= 4 && bytes[n - 4] == 13 && bytes[n - 3] == 10 && bytes[n - 2] == 13 && bytes[n - 1] == 10) break;
                }
                string[] lines = Encoding.ASCII.GetString(input.ToArray()).Split("\r\n");
                string[] line = lines[0].Split(' ');
                string path = line[1];
                var headers = lines.Skip(1).Where(value => value.Contains(':')).Select(value => value.Split(':', 2)).ToDictionary(pair => pair[0], pair => pair[1].Trim(), StringComparer.OrdinalIgnoreCase);
                int length = headers.TryGetValue("Content-Length", out string? size) ? int.Parse(size) : 0;
                if (length < 0 || length > 4 * 1024 * 1024) throw new IOException("Oversized fixture request.");
                var body = new byte[length];
                await stream.ReadExactlyAsync(body, _stop.Token);
                var received = new ReceivedRequest(line[0], path, headers, body);
                Requests.Enqueue(path);
                Received.Enqueue(received);
                Reply reply = RespondToRequest == null ? await _respond(path, _stop.Token) : await RespondToRequest(received, _stop.Token);
                string header = $"HTTP/1.1 {reply.Status} Test\r\nConnection: close\r\nContent-Type: {reply.ContentType}\r\n" +
                    (reply.Chunked ? "Transfer-Encoding: chunked\r\n" : $"Content-Length: {reply.Content.Length}\r\n") + reply.Headers + "\r\n";
                await stream.WriteAsync(Encoding.ASCII.GetBytes(header), _stop.Token);
                if (reply.Chunked) await stream.WriteAsync(Encoding.ASCII.GetBytes(reply.Content.Length.ToString("X") + "\r\n"), _stop.Token);
                if (line[0] != "HEAD") await stream.WriteAsync(reply.Content, _stop.Token);
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
    internal sealed record ReceivedRequest(string Method, string Path, IReadOnlyDictionary<string, string> Headers, byte[] Body);
}
