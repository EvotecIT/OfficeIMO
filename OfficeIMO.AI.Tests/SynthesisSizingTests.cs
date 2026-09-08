using System.Text.Json;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class SynthesisSizingTests {
    [Fact]
    public async Task GroupsRespectTransportLimitsWhenRequestIdsGainADigit() {
        var executor = new SizedExecutor();
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, new OfficeDocumentReadResult {
            Blocks = Enumerable.Range(0, 20).Select(index => new OfficeDocumentBlock { Text = "Fact " + index }).ToArray()
        });
        await new OfficeAiEngine(executor).RunAsync(document, new() { Operation = OfficeAiOperation.Summarize,
            Instruction = "Summarize", Limits = new() { MaxRequestCharacters = 4096, MaxRequests = 100 } });
        Assert.Contains(executor.Requests, request => request.RequestId.EndsWith("-summary-10", StringComparison.Ordinal));
        Assert.All(executor.Requests, request => Assert.InRange(executor.MeasureRequestCharacters(request), 1, 4096));
    }

    private sealed class SizedExecutor : IOfficeAiExecutor {
        public List<OfficeAiExecutionRequest> Requests { get; } = new();
        public OfficeAiExecutionProfile Profile { get; } = new() { Id = "sized", Provider = "fixture", Model = "fixture", IsLocal = true,
            MaxRequestCharacters = 4096 };
        public int MeasureRequestCharacters(OfficeAiExecutionRequest request) {
            using var json = JsonDocument.Parse(request.InputJson);
            // A transport with fixed-size draft slots plus the literal request ID in its envelope.
            if (json.RootElement.TryGetProperty("drafts", out var drafts))
                return checked(drafts.GetArrayLength() * 2000 + request.RequestId.Length + 54);
            return json.RootElement.GetProperty("evidence").GetArrayLength() > 1 ? 4097 : 1024;
        }
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            Requests.Add(request);
            using var json = JsonDocument.Parse(request.InputJson);
            if (json.RootElement.TryGetProperty("drafts", out var drafts))
                return Task.FromResult(new OfficeAiExecutionResponse(JsonSerializer.Serialize(new { claims = new[] {
                    new { text = "Combined facts", sourceClaimIds = drafts.EnumerateArray().Select(draft => draft.GetProperty("id").GetString()).ToArray() }
                } })));
            var evidence = json.RootElement.GetProperty("evidence")[0];
            return Task.FromResult(new OfficeAiExecutionResponse(JsonSerializer.Serialize(new {
                claims = new[] { new { text = evidence.GetProperty("text").GetString(), evidence = new[] {
                    new { id = evidence.GetProperty("id").GetString(), quote = evidence.GetProperty("text").GetString() } } } },
                fields = Array.Empty<object>(), blocks = Array.Empty<object>(), tables = Array.Empty<object>()
            })));
        }
    }
}
