using System.Collections.Concurrent;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using OfficeIMO.AI.IntelligenceX;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class CompatibleAdapterTests {
    private const string Answer = """
        {"status":"ok","claims":[{"text":"The total is 42.","evidence":[{"id":"e1","quote":"42"}]}],"fields":[],"blocks":[],"tables":[]}
        """;

    [Theory]
    [InlineData(true, true)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public async Task SameDocumentOperationWorksThroughLocalAndHostedCompatibleProfiles(bool local, bool schema) {
        using var server = new Server(context => Reply(context, Answer));
        using var executor = await IntelligenceXOfficeAiExecutor.ConnectAsync(Profile(local, schema), new() {
            Transport = OfficeAiIntelligenceXTransport.CompatibleHttp, Endpoint = server.Endpoint
        });
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(Document(), new() { Instruction = "What is the total?", AllowRemoteProcessing = !local });
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        Assert.Equal("The total is 42.", Assert.Single(result.Claims).Text);
        using var request = JsonDocument.Parse(Assert.Single(server.Requests).Body);
        Assert.Equal(schema, request.RootElement.TryGetProperty("response_format", out _));
        Assert.False(request.RootElement.TryGetProperty("tools", out _));
        Assert.Equal(!schema, result.Diagnostics.Contains("prompted-json-local-validation"));
    }

    [Fact]
    public async Task LocalEndpointDoesNotFollowRedirects() {
        using var server = new Server(context => {
            if (context.Request.Url!.AbsolutePath == "/v1/chat/completions") {
                context.Response.StatusCode = 307;
                context.Response.RedirectLocation = "/redirected";
                context.Response.Close(); return Task.CompletedTask;
            }
            return Reply(context, Answer);
        });
        using var executor = await IntelligenceXOfficeAiExecutor.ConnectAsync(Profile(), new() {
            Transport = OfficeAiIntelligenceXTransport.CompatibleHttp, Endpoint = server.Endpoint
        });
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(Document(), new() { Instruction = "Total?" });
        Assert.Equal(OfficeAiResultStatus.InvalidResponse, result.Status);
        Assert.Equal("/v1/chat/completions", Assert.Single(server.Requests).Path);
    }

    [Fact]
    public async Task FreshTreatmentRequestsDoNotReplayPriorDocumentText() {
        using var server = new Server(context => Reply(context, Answer));
        using var executor = await IntelligenceXOfficeAiExecutor.ConnectAsync(Profile(), new() {
            Transport = OfficeAiIntelligenceXTransport.CompatibleHttp, Endpoint = server.Endpoint
        });
        var engine = new OfficeAiEngine(executor);
        await engine.RunAsync(Document("prior-document-secret 42"), new() { Instruction = "Total?" });
        await engine.RunAsync(Document("current-document 42"), new() { Instruction = "Total?" });
        var calls = server.Requests.ToArray();
        Assert.Equal(2, calls.Length);
        Assert.DoesNotContain("prior-document-secret", calls[1].Body);
        Assert.Contains("current-document", calls[1].Body);
    }

    [Theory]
    [InlineData("https://example.invalid/v1")]
    [InlineData("http://user:secret@localhost:1234/v1")]
    [InlineData("http://localhost:1234/v1?key=secret")]
    public async Task InvalidLocalRoutesAreRejectedBeforeConnection(string endpoint) {
        await Assert.ThrowsAsync<ArgumentException>(() => IntelligenceXOfficeAiExecutor.ConnectAsync(Profile(), new() {
            Transport = OfficeAiIntelligenceXTransport.CompatibleHttp, Endpoint = new(endpoint)
        }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DeeplyNestedProviderContentAndEnvelopeFailWithoutTerminatingTheHost(bool envelope) {
        string deep = new string('[', 15_000) + "0" + new string(']', 15_000);
        using var server = new Server(async context => {
            if (!envelope) { await Reply(context, deep); return; }
            byte[] bytes = Encoding.UTF8.GetBytes("{\"unexpected\":" + deep + "}");
            context.Response.ContentType = "application/json"; context.Response.ContentLength64 = bytes.Length;
            await context.Response.OutputStream.WriteAsync(bytes); context.Response.Close();
        });
        using var executor = await IntelligenceXOfficeAiExecutor.ConnectAsync(Profile(), new() {
            Transport = OfficeAiIntelligenceXTransport.CompatibleHttp, Endpoint = server.Endpoint
        });
        var result = await new OfficeAiEngine(executor).RunAsync(Document(), new() { Instruction = "Total?" });
        Assert.Equal(OfficeAiResultStatus.InvalidResponse, result.Status);
        Assert.Empty(result.Claims);
        Assert.Single(server.Requests);
    }

    private static OfficeAiExecutionProfile Profile(bool local = true, bool schema = true) => new() {
        Id = "protocol-fixture", Provider = "fixture", Model = "fixture-model", IsLocal = local, EnforcesJsonSchema = schema
    };
    private static OfficeAiDocument Document(string text = "Total 42") => OfficeAiDocument.FromReadResult(Encoding.UTF8.GetBytes(text),
        new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Text = text } } });
    private static async Task Reply(HttpListenerContext context, string text) {
        byte[] bytes = JsonSerializer.SerializeToUtf8Bytes(new {
            choices = new[] { new { index = 0, message = new { role = "assistant", content = text }, finish_reason = "stop" } },
            usage = new { prompt_tokens = 10, completion_tokens = 10, total_tokens = 20 }
        });
        context.Response.ContentType = "application/json"; context.Response.ContentLength64 = bytes.Length;
        await context.Response.OutputStream.WriteAsync(bytes); context.Response.Close();
    }

    private sealed class Server : IDisposable {
        private readonly HttpListener _listener = new();
        private readonly Task _pending;
        public Uri Endpoint { get; }
        public ConcurrentQueue<(string Path, string Body)> Requests { get; } = new();
        public Server(Func<HttpListenerContext, Task> respond) {
            using var reservation = new TcpListener(IPAddress.Loopback, 0); reservation.Start();
            int port = ((IPEndPoint)reservation.LocalEndpoint).Port; reservation.Stop();
            Endpoint = new Uri($"http://127.0.0.1:{port}/v1/");
            _listener.Prefixes.Add($"http://127.0.0.1:{port}/"); _listener.Start();
            _pending = RunAsync(respond);
        }
        private async Task RunAsync(Func<HttpListenerContext, Task> respond) {
            try {
                while (_listener.IsListening) {
                    HttpListenerContext context = await _listener.GetContextAsync();
                    using var reader = new StreamReader(context.Request.InputStream);
                    Requests.Enqueue((context.Request.Url!.AbsolutePath, await reader.ReadToEndAsync()));
                    await respond(context);
                }
            } catch (HttpListenerException) when (!_listener.IsListening) { }
              catch (ObjectDisposedException) { }
        }
        public void Dispose() { _listener.Close(); _pending.GetAwaiter().GetResult(); }
    }
}
