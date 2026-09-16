using System.Collections.Concurrent;
using System.Net;
using System.Net.Sockets;
using System.Text;
using OfficeIMO.Html.Runtime;

internal sealed record ControlledAcquisitionCorpus(
    IReadOnlyList<ProbeCase> RenderCases,
    IReadOnlyList<AcquisitionProbeResult> Results) {

    internal static async Task<ControlledAcquisitionCorpus> CreateAsync() {
        var renderCases = new List<ProbeCase>();
        var results = new List<AcquisitionProbeResult>();
        await AddSameHostRedirectAsync(renderCases, results);
        await AddCrossHostRedirectAsync(renderCases, results);
        await AddDnsRebindingRejectionAsync(results);
        await AddOversizedResponseRejectionAsync(results);
        return new ControlledAcquisitionCorpus(renderCases.AsReadOnly(), results.AsReadOnly());
    }

    private static async Task AddSameHostRedirectAsync(List<ProbeCase> renderCases, List<AcquisitionProbeResult> results) {
        const string name = "acquisition-same-host-redirect";
        await using var server = new ControlledHttpServer(path => path switch {
            "/start" => ControlledHttpReply.Redirect("/final"),
            "/final" => ControlledHttpReply.Html(Page("Same-host redirect ready")),
            _ => ControlledHttpReply.NotFound()
        });
        var resolutions = new List<string>();
        var connections = new List<string>();
        var broker = Broker(server, ["page.example.test"], (host, _) => {
            resolutions.Add(host + "=93.184.216.34");
            return Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") });
        }, connections);
        try {
            Uri requested = new("http://page.example.test/start");
            HtmlPublicResourceResult acquired = await broker.FetchAsync(requested);
            Uri final = new("http://page.example.test/final");
            Require(acquired.Resource.Url == requested && acquired.Resource.FinalUrl == final &&
                acquired.Resource.StatusCode == 200 && acquired.ConnectedAddress.Equals(IPAddress.Parse("93.184.216.34")),
                "wrong final response provenance");
            RequireRedirect(acquired, requested, final, IPAddress.Parse("93.184.216.34"));
            Require(resolutions.SequenceEqual(new[] {
                "page.example.test=93.184.216.34", "page.example.test=93.184.216.34"
            }, StringComparer.Ordinal), "each same-host redirect hop was not independently resolved");
            Require(connections.SequenceEqual(new[] { "93.184.216.34:80", "93.184.216.34:80" }, StringComparer.Ordinal),
                "connector did not receive the validated same-host addresses");
            Require(server.Requests.SequenceEqual(new[] {
                "http://page.example.test/start", "http://page.example.test/final"
            }, StringComparer.Ordinal), "wrong same-host HTTP request authorities or paths");
            renderCases.Add(RenderCase(name, acquired));
            results.Add(Result(name, true, requested, acquired, resolutions, connections, server.Requests, null));
        } catch (Exception error) {
            results.Add(Result(name, false, new Uri("http://page.example.test/start"), null,
                resolutions, connections, server.Requests, error));
        }
    }

    private static async Task AddCrossHostRedirectAsync(List<ProbeCase> renderCases, List<AcquisitionProbeResult> results) {
        const string name = "acquisition-cross-host-redirect";
        await using var server = new ControlledHttpServer(path => path switch {
            "/start" => ControlledHttpReply.Redirect("http://assets.example.test/final"),
            "/final" => ControlledHttpReply.Html(Page("Cross-host redirect ready")),
            _ => ControlledHttpReply.NotFound()
        });
        var resolutions = new List<string>();
        var connections = new List<string>();
        var broker = Broker(server, ["page.example.test", "assets.example.test"], (host, _) => {
            string address = host.Equals("assets.example.test", StringComparison.Ordinal) ? "1.1.1.1" : "93.184.216.34";
            resolutions.Add(host + "=" + address);
            return Task.FromResult(new[] { IPAddress.Parse(address) });
        }, connections);
        try {
            Uri requested = new("http://page.example.test/start");
            HtmlPublicResourceResult acquired = await broker.FetchAsync(requested);
            Uri final = new("http://assets.example.test/final");
            Require(acquired.Resource.Url == requested && acquired.Resource.FinalUrl == final &&
                acquired.Resource.StatusCode == 200 && acquired.ConnectedAddress.Equals(IPAddress.Parse("1.1.1.1")),
                "wrong final response provenance");
            RequireRedirect(acquired, requested, final, IPAddress.Parse("93.184.216.34"));
            Require(resolutions.SequenceEqual(new[] {
                "page.example.test=93.184.216.34", "assets.example.test=1.1.1.1"
            }, StringComparer.Ordinal), "redirect host was not independently resolved");
            Require(connections.SequenceEqual(new[] { "93.184.216.34:80", "1.1.1.1:80" }, StringComparer.Ordinal),
                "connector did not receive the validated per-hop addresses");
            Require(server.Requests.SequenceEqual(new[] {
                "http://page.example.test/start", "http://assets.example.test/final"
            }, StringComparer.Ordinal), "wrong cross-host HTTP request authorities or paths");
            renderCases.Add(RenderCase(name, acquired));
            results.Add(Result(name, true, requested, acquired, resolutions, connections, server.Requests, null));
        } catch (Exception error) {
            results.Add(Result(name, false, new Uri("http://page.example.test/start"), null,
                resolutions, connections, server.Requests, error));
        }
    }

    private static async Task AddDnsRebindingRejectionAsync(List<AcquisitionProbeResult> results) {
        const string name = "acquisition-dns-rebinding-rejected";
        await using var server = new ControlledHttpServer(path => path == "/start"
            ? ControlledHttpReply.Redirect("/final") : ControlledHttpReply.Html(Page("must not render")));
        var resolutions = new List<string>();
        var connections = new List<string>();
        int count = 0;
        var broker = Broker(server, ["page.example.test"], (host, _) => {
            string address = Interlocked.Increment(ref count) == 1 ? "93.184.216.34" : "127.0.0.1";
            resolutions.Add(host + "=" + address);
            return Task.FromResult(new[] { IPAddress.Parse(address) });
        }, connections);
        Exception? rejection = null;
        try { await broker.FetchAsync(new Uri("http://page.example.test/start")); }
        catch (Exception error) { rejection = error; }
        bool passed = rejection?.ToString().Contains("non-public address", StringComparison.Ordinal) == true &&
            connections.SequenceEqual(new[] { "93.184.216.34:80" }, StringComparer.Ordinal) &&
            server.Requests.SequenceEqual(new[] { "http://page.example.test/start" }, StringComparer.Ordinal);
        results.Add(Result(name, passed, new Uri("http://page.example.test/start"), null,
            resolutions, connections, server.Requests,
            passed ? rejection : rejection ?? new IOException("DNS rebinding was not rejected.")));
    }

    private static async Task AddOversizedResponseRejectionAsync(List<AcquisitionProbeResult> results) {
        const string name = "acquisition-oversized-response-rejected";
        await using var server = new ControlledHttpServer(_ => new ControlledHttpReply([], "text/html; charset=utf-8",
            200, DeclaredLength: 4 * 1024 * 1024 + 1));
        var resolutions = new List<string>();
        var connections = new List<string>();
        var broker = Broker(server, ["page.example.test"], (host, _) => {
            resolutions.Add(host + "=93.184.216.34");
            return Task.FromResult(new[] { IPAddress.Parse("93.184.216.34") });
        }, connections);
        Exception? rejection = null;
        try { await broker.FetchAsync(new Uri("http://page.example.test/large")); }
        catch (Exception error) { rejection = error; }
        bool passed = rejection?.ToString().Contains("byte budget", StringComparison.Ordinal) == true &&
            connections.SequenceEqual(new[] { "93.184.216.34:80" }, StringComparer.Ordinal) &&
            server.Requests.SequenceEqual(new[] { "http://page.example.test/large" }, StringComparer.Ordinal);
        results.Add(Result(name, passed, new Uri("http://page.example.test/large"), null,
            resolutions, connections, server.Requests,
            passed ? rejection : rejection ?? new IOException("Oversized response was not rejected.")));
    }

    private static ProbeCase RenderCase(string name, HtmlPublicResourceResult acquired) {
        string text = name.Contains("same-host", StringComparison.Ordinal)
            ? "Same-host redirect ready" : "Cross-host redirect ready";
        return new ProbeCase(name, HtmlPublicResourceBroker.DecodeUtf8Html(acquired.Resource),
            $"document.querySelector('#result')?.textContent === '{text}'", 8 * 1024 * 1024,
            ExpectedVisibleText: text, ExpectBlueInk: true, ExpectedDiscoveryRounds: [],
            DocumentUrl: acquired.Resource.FinalUrl);
    }

    private static string Page(string text) =>
        $"<!doctype html><style>body{{font:16px sans-serif}}#result{{color:#0055aa}}</style><p id=result>{text}</p>";

    private static HtmlPublicResourceBroker Broker(ControlledHttpServer server, IEnumerable<string> hosts,
        Func<string, CancellationToken, Task<IPAddress[]>> resolve, List<string> connections) =>
        new(hosts, resolve, async (address, port, token) => {
            connections.Add(address + ":" + port);
            var socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            try {
                await socket.ConnectAsync(IPAddress.Loopback, server.Port, token);
                return new NetworkStream(socket, ownsSocket: true);
            } catch { socket.Dispose(); throw; }
        });

    private static AcquisitionProbeResult Result(string name, bool passed, Uri requestedUrl, HtmlPublicResourceResult? acquired,
        IReadOnlyList<string> resolutions, IReadOnlyList<string> connections, IEnumerable<string> requests, Exception? error) =>
        new(name, passed, requestedUrl.AbsoluteUri, acquired?.Resource.FinalUrl.AbsoluteUri,
            acquired?.ConnectedAddress.ToString(), acquired?.Redirects.Select(hop => new AcquisitionRedirect(
                hop.From.AbsoluteUri, hop.To.AbsoluteUri, hop.StatusCode, hop.ConnectedAddress.ToString())).ToArray() ?? [],
            resolutions.ToArray(), connections.ToArray(), requests.ToArray(), error?.GetType().Name, error?.Message);

    private static void Require(bool condition, string message) {
        if (!condition) throw new IOException(message);
    }

    private static void RequireRedirect(HtmlPublicResourceResult acquired, Uri from, Uri to,
        IPAddress connectedAddress) {
        Require(acquired.Redirects.Count == 1, "wrong redirect count");
        HtmlPublicRedirect redirect = acquired.Redirects[0];
        Require(redirect.From == from && redirect.To == to && redirect.StatusCode == 302 &&
            redirect.ConnectedAddress.Equals(connectedAddress), "wrong redirect provenance");
    }
}

internal sealed record AcquisitionProbeResult(string Name, bool Passed, string? RequestedUrl, string? FinalUrl,
    string? ConnectedAddress, AcquisitionRedirect[] Redirects, string[] Resolutions, string[] Connections,
    string[] ServerRequests, string? ErrorKind, string? Error);
internal sealed record AcquisitionRedirect(string From, string To, int StatusCode, string ConnectedAddress);

internal sealed class ControlledHttpServer : IAsyncDisposable {
    private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
    private readonly CancellationTokenSource _stop = new();
    private readonly Func<string, ControlledHttpReply> _respond;
    private readonly List<Task> _clients = [];
    private readonly Task _accept;
    internal ConcurrentQueue<string> Requests { get; } = new();
    internal int Port => ((IPEndPoint)_listener.LocalEndpoint).Port;

    internal ControlledHttpServer(Func<string, ControlledHttpReply> respond) {
        _respond = respond;
        _listener.Start();
        _accept = AcceptAsync();
    }

    private async Task AcceptAsync() {
        try {
            while (!_stop.IsCancellationRequested)
                _clients.Add(ServeAsync(await _listener.AcceptTcpClientAsync(_stop.Token)));
        } catch (OperationCanceledException) when (_stop.IsCancellationRequested) { }
    }

    private async Task ServeAsync(TcpClient client) {
        using (client) {
            try {
                NetworkStream stream = client.GetStream();
                using var input = new MemoryStream();
                var one = new byte[1];
                while (input.Length < 32768) {
                    if (await stream.ReadAsync(one, _stop.Token) != 1) throw new IOException("No request.");
                    input.WriteByte(one[0]);
                    byte[] buffer = input.GetBuffer();
                    int length = (int)input.Length;
                    if (length >= 4 && buffer[length - 4] == 13 && buffer[length - 3] == 10 &&
                        buffer[length - 2] == 13 && buffer[length - 1] == 10) break;
                }
                string[] lines = Encoding.ASCII.GetString(input.ToArray()).Split("\r\n");
                string requestLine = lines[0];
                string path = requestLine.Split(' ')[1];
                string? host = lines.Skip(1).Select(line => line.Split(':', 2))
                    .Where(pair => pair.Length == 2 && pair[0].Equals("Host", StringComparison.OrdinalIgnoreCase))
                    .Select(pair => pair[1].Trim()).SingleOrDefault();
                if (string.IsNullOrEmpty(host)) throw new IOException("Fixture request omitted its Host authority.");
                Requests.Enqueue("http://" + host + path);
                ControlledHttpReply reply = _respond(path);
                int declaredLength = reply.DeclaredLength ?? reply.Content.Length;
                string location = reply.Location == null ? string.Empty : "Location: " + reply.Location + "\r\n";
                string header = $"HTTP/1.1 {reply.Status} Test\r\nConnection: close\r\nContent-Type: {reply.ContentType}\r\nContent-Length: {declaredLength}\r\n{location}\r\n";
                await stream.WriteAsync(Encoding.ASCII.GetBytes(header), _stop.Token);
                if (reply.Content.Length != 0) await stream.WriteAsync(reply.Content, _stop.Token);
            } catch (OperationCanceledException) when (_stop.IsCancellationRequested) { }
            catch (IOException) { }
        }
    }

    public async ValueTask DisposeAsync() {
        _stop.Cancel();
        _listener.Stop();
        await _accept;
        await Task.WhenAll(_clients);
        _stop.Dispose();
    }
}

internal sealed record ControlledHttpReply(byte[] Content, string ContentType, int Status,
    string? Location = null, int? DeclaredLength = null) {
    internal static ControlledHttpReply Html(string html) =>
        new(Encoding.UTF8.GetBytes(html), "text/html; charset=utf-8", 200);
    internal static ControlledHttpReply Redirect(string location) =>
        new([], "text/plain", 302, location);
    internal static ControlledHttpReply NotFound() =>
        new(Encoding.UTF8.GetBytes("not found"), "text/plain", 404);
}
