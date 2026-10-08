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
        Assert.Contains("synthesis-no-progress", result.Diagnostics);
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
    public async Task MultiBatchQuestionsCombineObservationsOrReportInsufficientEvidence(OfficeAiOperation operation, bool empty) {
        var executor = new Executor { EmptyClaims = empty };
        var request = Request() with { Operation = operation };
        var result = await new OfficeAiEngine(executor).RunAsync(Document(new string('a', 110000)), request);
        Assert.True(result.RequestCount > 1);
        Assert.Equal(empty ? OfficeAiResultStatus.InsufficientEvidence : OfficeAiResultStatus.Completed, result.Status);
        Assert.Equal(empty ? OfficeAiSynthesisStatus.NotRequired : OfficeAiSynthesisStatus.Completed, result.SynthesisStatus);
        Assert.Empty(result.OmittedEvidenceIds);
        var single = await new OfficeAiEngine(new Executor { EmptyClaims = empty }).RunAsync(Document("one batch"), request);
        Assert.Equal(empty ? OfficeAiResultStatus.InsufficientEvidence : OfficeAiResultStatus.Completed, single.Status);
        Assert.Equal(OfficeAiSynthesisStatus.NotRequired, single.SynthesisStatus);
    }

    [Fact]
    public async Task OversizedEscapedUnicodeEvidenceRetainsEveryCharacterAndOriginalQuoteOffsets() {
        string source = string.Concat(Enumerable.Repeat("Line \"quoted\" 😀 with tab\tand newline\n", 2000));
        var executor = new Executor();
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(Document(source), Request());
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        Assert.Equal(OfficeAiSynthesisStatus.Completed, result.SynthesisStatus);
        Assert.Equal(new[] { "e1" }, result.ProcessedEvidenceIds);
        Assert.Empty(result.OmittedEvidenceIds);
        Assert.True(result.RequestCount > 1);
        Assert.Equal(executor.Requests.Count, result.RequestCount);
        var fragments = executor.Requests.SelectMany(request => {
            using var json = JsonDocument.Parse(request.InputJson);
            if (!json.RootElement.TryGetProperty("evidence", out var evidence)) return Array.Empty<string>();
            return evidence.EnumerateArray().Select(item => item.GetProperty("text").GetString()!).ToArray();
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
    [InlineData(OfficeAiOperation.Summarize, "valid", OfficeAiSynthesisStatus.Completed)]
    [InlineData(OfficeAiOperation.Summarize, "unknown-id", OfficeAiSynthesisStatus.Incomplete)]
    [InlineData(OfficeAiOperation.Summarize, "missing-source", OfficeAiSynthesisStatus.Incomplete)]
    [InlineData(OfficeAiOperation.Summarize, "truncated", OfficeAiSynthesisStatus.Incomplete)]
    [InlineData(OfficeAiOperation.Ask, "valid", OfficeAiSynthesisStatus.Completed)]
    [InlineData(OfficeAiOperation.Ask, "unknown-id", OfficeAiSynthesisStatus.Incomplete)]
    [InlineData(OfficeAiOperation.Ask, "missing-source", OfficeAiSynthesisStatus.Incomplete)]
    [InlineData(OfficeAiOperation.Ask, "truncated", OfficeAiSynthesisStatus.Incomplete)]
    [InlineData(OfficeAiOperation.Explain, "valid", OfficeAiSynthesisStatus.Completed)]
    public async Task WholeDocumentReasoningPreservesCitationsAndRejectsInventedOrDroppedSourceClaims(OfficeAiOperation operation, string mode, OfficeAiSynthesisStatus expected) {
        var executor = new Executor { SynthesisMode = mode, ExpectedOperation = operation };
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(
            Document("North total 42. " + new string('x', 30000), "South total 57. " + new string('y', 30000)),
            Request() with { Operation = operation });
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
            Assert.Contains(operation == OfficeAiOperation.Summarize ? "summary-synthesis-incomplete" : "answer-synthesis-incomplete", result.Diagnostics);
        }
    }

    [Theory]
    [InlineData(OfficeAiOperation.Summarize)]
    [InlineData(OfficeAiOperation.Ask)]
    public async Task SynthesisSharesRequestBudgetAndKeepsDraftsWhenNoCallsRemain(OfficeAiOperation operation) {
        var executor = new Executor();
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(
            Document(new string('a', 30000), new string('b', 30000)),
            Request() with { Operation = operation, Limits = new() { MaxRequests = 2 } });
        Assert.Equal(OfficeAiSynthesisStatus.Incomplete, result.SynthesisStatus);
        Assert.Equal(2, result.RequestCount);
        Assert.Equal(2, result.Claims.Count);
        Assert.Empty(result.OmittedEvidenceIds);
        Assert.Contains("synthesis-request-budget-exceeded", result.Diagnostics);
    }

    [Theory]
    [InlineData(OfficeAiOperation.Ask)]
    [InlineData(OfficeAiOperation.Explain)]
    [InlineData(OfficeAiOperation.Summarize)]
    public async Task DefaultReserveCombinesProcessedEvidenceAndDisclosesUnseenRemainder(OfficeAiOperation operation) {
        var executor = new Executor();
        var result = await new OfficeAiEngine(executor).RunAsync(
            Document(new string('a', 30000), new string('b', 30000), new string('c', 30000)),
            Request() with { Operation = operation, Limits = new() { MaxRequests = 3 } });
        Assert.Equal(OfficeAiSynthesisStatus.Completed, result.SynthesisStatus);
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Equal(3, result.RequestCount);
        Assert.Equal(new[] { "e1", "e2" }, result.ProcessedEvidenceIds);
        Assert.Equal(new[] { "e3" }, result.OmittedEvidenceIds);
        Assert.Contains("evidence-budget-exceeded", result.Diagnostics);
        using var synthesis = JsonDocument.Parse(executor.Requests[^1].InputJson);
        var coverage = synthesis.RootElement.GetProperty("coverage");
        Assert.Equal(1, coverage.GetProperty("omittedEvidenceCount").GetInt32());
        Assert.Empty(coverage.GetProperty("emptyPages").EnumerateArray());
        Assert.False(coverage.GetProperty("sourceReaderDiagnostics").GetBoolean());
        Assert.Equal(new[] { "e1", "e2" }, Assert.Single(result.Claims).Citations.Select(c => c.EvidenceId));
    }

    [Fact]
    public async Task CallerCanPreferEvidenceCoverageOverSynthesisReservation() {
        var executor = new Executor();
        var result = await new OfficeAiEngine(executor).RunAsync(
            Document(new string('a', 30000), new string('b', 30000), new string('c', 30000)),
            Request() with { Limits = new() { MaxRequests = 3, ReservedSynthesisRequests = 0 } });
        Assert.Equal(OfficeAiSynthesisStatus.Incomplete, result.SynthesisStatus);
        Assert.Equal(3, result.RequestCount);
        Assert.Empty(result.OmittedEvidenceIds);
        Assert.Equal(3, result.Claims.Count);
        Assert.Contains("synthesis-request-budget-exceeded", result.Diagnostics);
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
        public OfficeAiOperation? ExpectedOperation { get; init; }
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
                if (ExpectedOperation.HasValue) Assert.Equal(ExpectedOperation.Value.ToString(), json.RootElement.GetProperty("operation").GetString());
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
                output = JsonSerializer.Serialize(new { claims, fields = Array.Empty<object>(), blocks = Array.Empty<object>(), tables = Array.Empty<object>() });
            }
            return Task.FromResult(new OfficeAiExecutionResponse(output, IsComplete: !(json.RootElement.TryGetProperty("drafts", out _) && SynthesisMode == "truncated"), InputTokens: 1, OutputTokens: 2));
        }
    }
}
