using System.Text.Json;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class SynthesisSizingTests {
    [Theory]
    [InlineData(false, "exception")]
    [InlineData(true, "exception")]
    [InlineData(false, "negative")]
    [InlineData(true, "negative")]
    public async Task MeasurementFailureRetainsDraftsAndSanitizesPlanningErrors(bool duringSynthesis, string failure) {
        var executor = new SizedExecutor { FailMeasurementAt = duringSynthesis ? "synthesis" : "planning", MeasurementFailure = failure };
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, new OfficeDocumentReadResult {
            Blocks = new[] { new OfficeDocumentBlock { Text = "Alpha" }, new OfficeDocumentBlock { Text = "Beta" } }
        });
        var task = new OfficeAiEngine(executor).RunAsync(document, new() { Operation = OfficeAiOperation.Summarize,
            Instruction = "Summarize", Limits = new() { MaxRequestCharacters = 4096 } });
        if (duringSynthesis) {
            var result = await task;
            Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
            Assert.Equal(OfficeAiSynthesisStatus.Incomplete, result.SynthesisStatus);
            Assert.Equal(2, result.Claims.Count);
            Assert.Equal(2, result.RequestCount);
            Assert.DoesNotContain("sensitive", JsonSerializer.Serialize(result), StringComparison.Ordinal);
        } else {
            var exception = await Assert.ThrowsAsync<InvalidDataException>(() => task);
            Assert.DoesNotContain("sensitive", exception.ToString(), StringComparison.Ordinal);
            Assert.Null(exception.InnerException);
            Assert.Empty(executor.Requests);
        }
    }

    [Theory]
    [InlineData(false, "fatal")]
    [InlineData(true, "fatal")]
    [InlineData(false, "canceled")]
    [InlineData(true, "canceled")]
    public async Task MeasurementPreservesResourceFailureAndCancellation(bool duringSynthesis, string failure) {
        using var cancellation = new CancellationTokenSource();
        var executor = new SizedExecutor { FailMeasurementAt = duringSynthesis ? "synthesis" : "planning",
            MeasurementFailure = failure, Cancellation = cancellation };
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, new OfficeDocumentReadResult {
            Blocks = new[] { new OfficeDocumentBlock { Text = "Alpha" }, new OfficeDocumentBlock { Text = "Beta" } }
        });
        var task = new OfficeAiEngine(executor).RunAsync(document, new() { Operation = OfficeAiOperation.Summarize,
            Instruction = "Summarize", Limits = new() { MaxRequestCharacters = 4096 } }, cancellationToken: cancellation.Token);
        if (failure == "fatal") await Assert.ThrowsAsync<OutOfMemoryException>(() => task);
        else await Assert.ThrowsAnyAsync<OperationCanceledException>(() => task);
    }

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
        public string? FailMeasurementAt { get; init; }
        public string MeasurementFailure { get; init; } = "exception";
        public CancellationTokenSource? Cancellation { get; init; }
        public List<OfficeAiExecutionRequest> Requests { get; } = new();
        public OfficeAiExecutionProfile Profile { get; } = new() { Id = "sized", Provider = "fixture", Model = "fixture", IsLocal = true,
            MaxRequestCharacters = 4096 };
        public int MeasureRequestCharacters(OfficeAiExecutionRequest request) {
            using var json = JsonDocument.Parse(request.InputJson);
            if (FailMeasurementAt == (json.RootElement.TryGetProperty("drafts", out _) ? "synthesis" : "planning")) {
                if (MeasurementFailure == "negative") return -1;
                if (MeasurementFailure == "fatal") throw new OutOfMemoryException("fixture resource failure");
                if (MeasurementFailure == "canceled") { Cancellation!.Cancel(); return 0; }
                throw new InvalidOperationException("sensitive transport diagnostic");
            }
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
