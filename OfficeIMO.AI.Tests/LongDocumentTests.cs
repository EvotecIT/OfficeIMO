using System.Text;
using System.Text.Json;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class LongDocumentTests {
    [Theory]
    [InlineData("negative-usage")]
    [InlineData("invalid-executor")]
    public async Task InvalidSynthesisUsageMakesOperationTotalsUnknown(string mode) {
        var result = await new OfficeAiEngine(new Executor { SynthesisMode = mode }).RunAsync(
            Document("North total 42. " + new string('x', 30000), "South total 57. " + new string('y', 30000)),
            Request() with { Operation = OfficeAiOperation.Summarize });
        Assert.Equal(3, result.RequestCount);
        Assert.Equal(OfficeAiSynthesisStatus.Incomplete, result.SynthesisStatus);
        Assert.Null(result.InputTokens);
        Assert.Null(result.OutputTokens);
    }

    [Fact]
    public async Task SummaryStopsAfterTheFirstNonReducingPassAndRetainsSupportedDrafts() {
        var executor = new Executor { LargeDrafts = true, SynthesisMode = "shrink-then-stall" };
        var result = await new OfficeAiEngine(executor).RunAsync(Document(new string('a', 50000)), Request() with {
            Operation = OfficeAiOperation.Summarize,
            Limits = new() { MaxRequestCharacters = 12000, MaxRequests = 64, MaxSynthesisPasses = 5 }
        });
        int initialDrafts = 0, stalledDrafts = 0;
        foreach (var request in executor.Requests) {
            using var json = JsonDocument.Parse(request.InputJson);
            if (!json.RootElement.TryGetProperty("drafts", out var drafts)) continue;
            foreach (var draft in drafts.EnumerateArray()) {
                if (draft.GetProperty("text").GetString()!.Length > 4000) initialDrafts++;
                else stalledDrafts++;
            }
        }
        Assert.True(initialDrafts > 2);
        Assert.Equal(initialDrafts, stalledDrafts);
        Assert.Equal(initialDrafts, result.Claims.Count);
        Assert.All(result.Claims, claim => Assert.NotEmpty(claim.Citations));
        Assert.Equal(OfficeAiSynthesisStatus.Incomplete, result.SynthesisStatus);
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Empty(result.OmittedEvidenceIds);
    }

    [Fact]
    public async Task SynthesisOutOfMemoryEscapesWithoutAnotherModelCall() {
        var executor = new Executor { FatalOnCall = 3 };
        await Assert.ThrowsAsync<OutOfMemoryException>(() => new OfficeAiEngine(executor).RunAsync(
            Document("North total 42. " + new string('x', 30000), "South total 57. " + new string('y', 30000)),
            Request() with { Operation = OfficeAiOperation.Summarize }));
        Assert.Equal(3, executor.Requests.Count);
    }

    [Theory]
    [InlineData(OfficeAiOperation.Ask, false)]
    [InlineData(OfficeAiOperation.Ask, true)]
    [InlineData(OfficeAiOperation.Explain, false)]
    [InlineData(OfficeAiOperation.Explain, true)]
    public async Task MultiBatchQuestionsReportTheCrossBatchReasoningLimit(OfficeAiOperation operation, bool empty) {
        var executor = new Executor { EmptyClaims = empty };
        var request = Request() with { Operation = operation };
        var result = await new OfficeAiEngine(executor).RunAsync(Document(new string('a', 110000)), request);
        Assert.True(result.RequestCount > 1);
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Contains("cross-batch-reasoning-not-supported", result.Diagnostics);
        Assert.Empty(result.OmittedEvidenceIds);
        var single = await new OfficeAiEngine(new Executor { EmptyClaims = empty }).RunAsync(Document("one batch"), request);
        Assert.Equal(empty ? OfficeAiResultStatus.InsufficientEvidence : OfficeAiResultStatus.Completed, single.Status);
        Assert.DoesNotContain("cross-batch-reasoning-not-supported", single.Diagnostics);
    }

    [Fact]
    public async Task OversizedEscapedUnicodeEvidenceRetainsEveryCharacterAndOriginalQuoteOffsets() {
        string source = string.Concat(Enumerable.Repeat("Line \"quoted\" 😀 with tab\tand newline\n", 2000));
        var executor = new Executor();
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(Document(source), Request());
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Contains("cross-batch-reasoning-not-supported", result.Diagnostics);
        Assert.Equal(new[] { "e1" }, result.ProcessedEvidenceIds);
        Assert.Empty(result.OmittedEvidenceIds);
        Assert.True(result.RequestCount > 1);
        Assert.Equal(executor.Requests.Count, result.RequestCount);
        var fragments = executor.Requests.SelectMany(request => {
            using var json = JsonDocument.Parse(request.InputJson);
            return json.RootElement.GetProperty("evidence").EnumerateArray().Select(item => item.GetProperty("text").GetString()!).ToArray();
        }).ToArray();
        Assert.Equal(source, string.Concat(fragments));
        Assert.All(fragments, fragment => {
            Assert.False(char.IsLowSurrogate(fragment[0]));
            Assert.False(char.IsHighSurrogate(fragment[^1]));
        });
        Assert.Equal(source.Length, result.ProcessedTextRanges.Sum(range => range.Length));
        int offset = 0;
        foreach (OfficeAiEvidenceRange range in result.ProcessedTextRanges) {
            Assert.Equal(offset, range.Start); offset += range.Length;
        }
        foreach (OfficeAiCitation citation in result.Claims.SelectMany(claim => claim.Citations)) {
            Assert.Equal("e1", citation.EvidenceId);
            Assert.Equal(citation.Quote, source.Substring(citation.QuoteStart!.Value, citation.Quote!.Length));
        }
    }

    [Theory]
    [InlineData("valid", OfficeAiSynthesisStatus.Completed)]
    [InlineData("unknown-id", OfficeAiSynthesisStatus.Incomplete)]
    [InlineData("missing-source", OfficeAiSynthesisStatus.Incomplete)]
    [InlineData("truncated", OfficeAiSynthesisStatus.Incomplete)]
    public async Task WholeSummaryPreservesCitationsAndRejectsInventedOrDroppedSourceClaims(string mode, OfficeAiSynthesisStatus expected) {
        var executor = new Executor { SynthesisMode = mode };
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(
            Document("North total 42. " + new string('x', 30000), "South total 57. " + new string('y', 30000)),
            Request() with { Operation = OfficeAiOperation.Summarize });
        Assert.Equal(expected, result.SynthesisStatus);
        Assert.Equal(3, result.RequestCount);
        Assert.Equal(3, result.InputTokens);
        Assert.Equal(6, result.OutputTokens);
        Assert.Equal(new[] { "e1", "e2" }, result.ProcessedEvidenceIds);
        Assert.Empty(result.OmittedEvidenceIds);
        Assert.Equal(new[] { "e1", "e2" }, result.Claims.SelectMany(claim => claim.Citations).Select(citation => citation.EvidenceId).Distinct().ToArray());
        if (expected == OfficeAiSynthesisStatus.Completed) {
            Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
            Assert.Equal("Combined regional totals.", Assert.Single(result.Claims).Text);
        } else {
            Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
            Assert.Equal(2, result.Claims.Count);
            Assert.Contains("summary-synthesis-incomplete", result.Diagnostics);
        }
    }

    [Fact]
    public async Task SynthesisSharesRequestBudgetAndKeepsDraftsWhenNoCallsRemain() {
        var executor = new Executor();
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(
            Document(new string('a', 30000), new string('b', 30000)),
            Request() with { Operation = OfficeAiOperation.Summarize, Limits = new() { MaxRequests = 2 } });
        Assert.Equal(OfficeAiSynthesisStatus.Incomplete, result.SynthesisStatus);
        Assert.Equal(2, result.RequestCount);
        Assert.Equal(2, result.Claims.Count);
        Assert.Empty(result.OmittedEvidenceIds);
    }

    [Fact]
    public async Task HierarchicalSummaryReducesLargeDraftGroupsWithinSharedBudget() {
        var executor = new Executor { LargeDrafts = true };
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(Document(new string('a', 24000)),
            Request() with { Operation = OfficeAiOperation.Summarize, Limits = new() { MaxRequestCharacters = 12000 } });
        Assert.Equal(OfficeAiSynthesisStatus.Completed, result.SynthesisStatus);
        Assert.InRange(result.RequestCount, 5, 32);
        Assert.Equal("Combined regional totals.", Assert.Single(result.Claims).Text);
        Assert.Empty(result.OmittedEvidenceIds);
    }

    [Fact]
    public async Task FailedMiddleFragmentDoesNotEraseOtherCoverageOrClaimCompleteOriginalRecord() {
        var executor = new Executor { FailCall = 2 };
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(Document(new string('a', 110000)), Request());
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Empty(result.ProcessedEvidenceIds);
        Assert.Equal(new[] { "e1" }, result.OmittedEvidenceIds);
        Assert.Equal(2, result.ProcessedTextRanges.Count);
        Assert.True(result.ProcessedTextRanges[1].Start > result.ProcessedTextRanges[0].Length);
        Assert.Null(result.InputTokens);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(4)]
    public async Task SingleClaimProducingBatchNeedsNoSynthesisEvenWhenOtherBatchesAreEmpty(int budget) {
        var executor = new Executor { EmptyAfterFirstBatch = true };
        var result = await new OfficeAiEngine(executor).RunAsync(Document(new string('a', 30000), new string('b', 30000)),
            Request() with { Operation = OfficeAiOperation.Summarize, Limits = new() { MaxRequests = budget } });
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        Assert.Equal(OfficeAiSynthesisStatus.NotRequired, result.SynthesisStatus);
        Assert.Equal(2, result.RequestCount);
        Assert.Single(result.Claims);
        Assert.Equal(new[] { "e1", "e2" }, result.ProcessedEvidenceIds);
        Assert.Empty(result.OmittedEvidenceIds);
    }

    private static OfficeAiRequest Request() => new() { Instruction = "Report all totals." };
    private static OfficeAiDocument Document(params string[] texts) => OfficeAiDocument.FromReadResult(
        Encoding.UTF8.GetBytes(string.Join("\n", texts)), new OfficeDocumentReadResult {
            Blocks = texts.Select(text => new OfficeDocumentBlock { Kind = "paragraph", Text = text }).ToArray()
        });

    private sealed class Executor : IOfficeAiExecutor {
        public string SynthesisMode { get; init; } = "valid";
        public bool LargeDrafts { get; init; }
        public bool EmptyClaims { get; init; }
        public bool EmptyAfterFirstBatch { get; init; }
        public int FailCall { get; init; }
        public int FatalOnCall { get; init; }
        public OfficeAiExecutionProfile Profile { get; } = new() { Id = "bounded", Provider = "fixture", Model = "fixture", IsLocal = true };
        public List<OfficeAiExecutionRequest> Requests { get; } = new();
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            Assert.True(request.Instructions.Length + request.InputJson.Length + request.OutputSchema.Length <= Profile.MaxRequestCharacters);
            Requests.Add(request);
            if (Requests.Count == FatalOnCall) throw new OutOfMemoryException("fixture resource failure");
            if (Requests.Count == FailCall) throw new IOException("fixture transport failure");
            using var json = JsonDocument.Parse(request.InputJson);
            string output;
            if (json.RootElement.TryGetProperty("drafts", out var drafts)) {
                if (SynthesisMode == "negative-usage") return Task.FromResult(new OfficeAiExecutionResponse("{}", InputTokens: 1, OutputTokens: -1));
                if (SynthesisMode == "invalid-executor") throw new InvalidDataException("Invalid executor payload");
                if (SynthesisMode == "shrink-then-stall") {
                    output = JsonSerializer.Serialize(new { claims = drafts.EnumerateArray().Select(item => new {
                        text = item.GetProperty("text").GetString()![..4000],
                        sourceClaimIds = new[] { item.GetProperty("id").GetString()! }
                    }) });
                    return Task.FromResult(new OfficeAiExecutionResponse(output));
                }
                string[] ids = drafts.EnumerateArray().Select(item => item.GetProperty("id").GetString()!).ToArray();
                if (SynthesisMode == "unknown-id") ids[0] = "invented";
                if (SynthesisMode == "missing-source") ids = ids.Take(1).ToArray();
                output = JsonSerializer.Serialize(new { claims = new[] { new { text = "Combined regional totals.", sourceClaimIds = ids } } });
            } else {
                var claims = json.RootElement.GetProperty("evidence").EnumerateArray().Where(_ => !EmptyClaims && !(EmptyAfterFirstBatch && Requests.Count > 1)).Select(item => {
                    string text = item.GetProperty("text").GetString()!;
                    string quote = text[..Math.Min(12, text.Length)];
                    return new { text = quote + (LargeDrafts ? new string('z', 6000) : ""), evidence = new[] { new { id = item.GetProperty("id").GetString(), quote } } };
                }).ToArray();
                output = JsonSerializer.Serialize(new { status = "ok", claims, fields = Array.Empty<object>(), blocks = Array.Empty<object>(), tables = Array.Empty<object>() });
            }
            return Task.FromResult(new OfficeAiExecutionResponse(output, IsComplete: !(request.RequestId.Contains("summary") && SynthesisMode == "truncated"), InputTokens: 1, OutputTokens: 2));
        }
    }
}
