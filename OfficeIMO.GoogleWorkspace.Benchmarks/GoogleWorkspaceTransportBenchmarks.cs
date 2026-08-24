using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using System.Net;

namespace OfficeIMO.GoogleWorkspace.Benchmarks;

[MemoryDiagnoser]
[ShortRunJob(RuntimeMoniker.Net80)]
public class GoogleWorkspaceTransportBenchmarks {
    private byte[] _payload = Array.Empty<byte>();
    private HttpClient _client = null!;
    private GoogleWorkspaceHttpTransport _transport = null!;

    [Params(64 * 1024, 4 * 1024 * 1024)]
    public int PayloadBytes { get; set; }

    [GlobalSetup]
    public async Task Setup() {
        _payload = GoogleWorkspaceTransportScenario.CreatePayload(PayloadBytes);
        _transport = GoogleWorkspaceTransportScenario.CreateTransport(_payload, out _client);

        byte[] result = await GoogleWorkspaceTransportScenario.DownloadAsync(_transport, PayloadBytes).ConfigureAwait(false);
        if (!result.AsSpan().SequenceEqual(_payload)) {
            throw new InvalidOperationException("Transport preflight did not preserve the complete response payload.");
        }
    }

    [GlobalCleanup]
    public void Cleanup() {
        _transport.Dispose();
        _client.Dispose();
    }

    [Benchmark]
    public Task<byte[]> DownloadBytes() =>
        GoogleWorkspaceTransportScenario.DownloadAsync(_transport, PayloadBytes);
}

internal static class GoogleWorkspaceTransportScenario {
    public static byte[] CreatePayload(int length) {
        var payload = new byte[length];
        for (var index = 0; index < payload.Length; index++) {
            payload[index] = unchecked((byte)(index * 31 + 17));
        }
        return payload;
    }

    public static GoogleWorkspaceHttpTransport CreateTransport(byte[] payload, out HttpClient client) {
        client = new HttpClient(new PayloadHandler(payload));
        return new GoogleWorkspaceHttpTransport(new GoogleWorkspaceSessionOptions {
            HttpClient = client,
            MaxRetryCount = 0,
            RequestTimeout = TimeSpan.FromSeconds(30)
        });
    }

    public static Task<byte[]> DownloadAsync(GoogleWorkspaceHttpTransport transport, int payloadBytes) =>
        transport.SendBytesAsync(
            accessToken: string.Empty,
            method: HttpMethod.Get,
            uri: "https://www.googleapis.com/drive/v3/files/evidence?alt=media",
            requestSafety: GoogleWorkspaceRequestSafety.Safe,
            serviceName: "Google Drive API",
            report: new TranslationReport(),
            includeAuthorization: false,
            maxResponseBytes: payloadBytes);

    private sealed class PayloadHandler : HttpMessageHandler {
        private readonly byte[] _payload;

        public PayloadHandler(byte[] payload) {
            _payload = payload;
        }

        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new ByteArrayContent(_payload)
            });
    }
}
