using System.Text;
using System.Text.Json;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class EngineContractTests {
    private const string Empty = "{\"status\":\"insufficient\",\"claims\":[],\"fields\":[],\"blocks\":[],\"tables\":[]}";

    [Theory]
    [InlineData(OfficeAiOperation.Ask, "claims")]
    [InlineData(OfficeAiOperation.Explain, "claims")]
    [InlineData(OfficeAiOperation.Summarize, "claims")]
    [InlineData(OfficeAiOperation.ExtractFields, "fields")]
    [InlineData(OfficeAiOperation.Parse, "blocks,tables")]
    public async Task GenerationSchemaExcludesResultsFromOtherOperations(OfficeAiOperation operation, string enabledNames) {
        var executor = new Executor(Empty);
        await new OfficeAiEngine(executor).RunAsync(Document("Total 42"), Request() with {
            Operation = operation,
            Fields = operation == OfficeAiOperation.ExtractFields ? new[] { new OfficeAiFieldDefinition("total") } : Array.Empty<OfficeAiFieldDefinition>()
        });
        using JsonDocument schema = JsonDocument.Parse(Assert.Single(executor.Requests).OutputSchema);
        string[] enabled = enabledNames.Split(',');
        foreach (string name in new[] { "claims", "fields", "blocks", "tables" }) {
            JsonElement array = schema.RootElement.GetProperty("properties").GetProperty(name);
            Assert.Equal("array", array.GetProperty("type").GetString());
            Assert.Equal(enabled.Contains(name) ? 200 : 0, array.GetProperty("maxItems").GetInt32());
        }
    }

    [Fact]
    public async Task RemoteEvidenceRequiresConsentBeforeExecution() {
        var executor = new Executor(Empty, local: false);
        await Assert.ThrowsAsync<InvalidOperationException>(() => new OfficeAiEngine(executor).RunAsync(Document("Total 42"), Request()));
        Assert.Empty(executor.Requests);
    }

    [Theory]
    [InlineData("e1", "Total 42", true)]
    [InlineData("e2", "Total 42", false)]
    [InlineData("e1", "Total 99", false)]
    [InlineData("e1", null, false)]
    public async Task CitationMustIdentifyExactCapturedEvidence(string id, string? quote, bool valid) {
        var executor = new Executor(Claim(id, quote));
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(Document("Total 42"), Request());
        Assert.Equal(valid ? OfficeAiResultStatus.Completed : OfficeAiResultStatus.InvalidResponse, result.Status);
        if (valid) {
            Assert.True(Assert.Single(Assert.Single(result.Claims).Citations).QuoteMatched);
            Assert.Contains("semantic-support-not-assessed", result.Diagnostics);
        } else {
            Assert.Empty(result.Claims);
            Assert.Equal(new[] { "e1" }, result.OmittedEvidenceIds);
        }
    }

    [Fact]
    public async Task TruncatedAndDuplicatePropertyResponsesAreRejected() {
        foreach (OfficeAiExecutionResponse response in new[] {
            new OfficeAiExecutionResponse(Claim("e1", "Total 42"), IsComplete: false),
            new OfficeAiExecutionResponse(Empty.Replace("\"status\":", "\"status\":\"ok\",\"status\":"))
        }) {
            var executor = new Executor(response);
            OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(Document("Total 42"), Request());
            Assert.Equal(OfficeAiResultStatus.InvalidResponse, result.Status);
            Assert.Contains("invalid-provider-response", result.Diagnostics);
        }
    }

    [Theory]
    [InlineData("1 234,50", "pl-PL", "1234.50")]
    [InlineData("1,234.50", "en-US", "1234.50")]
    [InlineData("12,34,567.89", "hi-IN", "1234567.89")]
    public async Task FieldNormalizationUsesExplicitCulture(string raw, string culture, string expected) {
        var executor = new Executor(Field("amount", raw, "e1"));
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(Document(raw), Request() with {
            Operation = OfficeAiOperation.ExtractFields, Culture = culture,
            Fields = new[] { new OfficeAiFieldDefinition("amount", OfficeAiFieldType.Decimal) }
        });
        OfficeAiField field = Assert.Single(result.Fields);
        Assert.Equal(OfficeAiFieldStatus.Present, field.Status);
        Assert.Equal(expected, field.NormalizedValue);
    }

    [Fact]
    public async Task PageScopeAndSourceSnapshotRemainStableAfterReaderMutation() {
        var block = new OfficeDocumentBlock { Text = "Total 42", Location = new() { Page = 1 } };
        var source = new OfficeDocumentReadResult { Blocks = new[] { block, new OfficeDocumentBlock { Text = "Private page two", Location = new() { Page = 2 } } } };
        OfficeAiDocument document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, source);
        block.Text = "changed";
        var executor = new Executor(Claim("e1", "Total 42"));
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(document, Request() with { Pages = new[] { 1 } });
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        string input = Assert.Single(executor.Requests).InputJson;
        Assert.Contains("Total 42", input);
        Assert.DoesNotContain("Private page two", input);
        Assert.DoesNotContain("changed", input);
        Assert.Equal(new[] { "e1" }, result.ProcessedEvidenceIds);
    }

    [Fact]
    public async Task BudgetOmissionsRetainValidatedPartialCoverage() {
        var executor = new Executor(Empty);
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(Document(new string('x', 60_000)), Request() with { Limits = new() { MaxRequests = 1 } });
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Equal(new[] { "e1" }, result.OmittedEvidenceIds);
        Assert.Empty(result.ProcessedEvidenceIds);
        Assert.Single(executor.Requests);
        Assert.InRange(Assert.Single(result.ProcessedTextRanges).Length, 1, 59_999);
    }

    [Fact]
    public async Task CancellationDoesNotOverlapAnUncooperativeExecutor() {
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource<OfficeAiExecutionResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
        var executor = new Executor((_, _) => { started.TrySetResult(); return release.Task; });
        var engine = new OfficeAiEngine(executor);
        using var firstCancellation = new CancellationTokenSource();
        Task<OfficeAiResult> first = engine.RunAsync(Document("hello"), Request(), cancellationToken: firstCancellation.Token);
        await started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        firstCancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => first);
        using var secondCancellation = new CancellationTokenSource();
        Task<OfficeAiResult> second = engine.RunAsync(Document("hello"), Request(), cancellationToken: secondCancellation.Token);
        secondCancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => second);
        Assert.Single(executor.Requests);
        release.SetResult(new(Empty));
    }

    [Fact]
    public async Task ProviderExceptionsDoNotLeakCredentialsOrSource() {
        var executor = new Executor((_, _) => throw new InvalidOperationException("secret-token private-document"));
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(Document("hello"), Request());
        Assert.Equal(OfficeAiResultStatus.InvalidResponse, result.Status);
        Assert.DoesNotContain("secret-token", JsonSerializer.Serialize(result));
        Assert.Contains("provider-execution-failed", result.Diagnostics);
    }

    private static OfficeAiRequest Request() => new() { Instruction = "What is the total?" };
    [Theory]
    [InlineData("1,23")]
    [InlineData("1,,234")]
    public async Task MalformedLocaleGroupingDoesNotSilentlyChangeTheAmount(string raw) {
        OfficeAiResult result = await new OfficeAiEngine(new Executor(Field("amount", raw, "e1"))).RunAsync(Document(raw), Request() with {
            Operation = OfficeAiOperation.ExtractFields, Culture = "en-US", Fields = new[] { new OfficeAiFieldDefinition("amount", OfficeAiFieldType.Decimal) }
        });
        OfficeAiField field = Assert.Single(result.Fields);
        Assert.Equal(OfficeAiFieldStatus.Invalid, field.Status);
        Assert.Null(field.NormalizedValue);
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
    }
    [Fact]
    public async Task ConflictingValuesAcrossBatchesAreNotFlattenedIntoOneAnswer() {
        var reader = new OfficeDocumentReadResult { Blocks = new[] {
            new OfficeDocumentBlock { Text = "42 " + new string('a', 30_000), Location = new() { Page = 1 } },
            new OfficeDocumentBlock { Text = "45 " + new string('b', 30_000), Location = new() { Page = 2 } }
        } };
        OfficeAiDocument document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, reader);
        var executor = new Executor((input, _) => {
            using var json = JsonDocument.Parse(input.InputJson);
            string id = json.RootElement.GetProperty("evidence")[0].GetProperty("id").GetString()!;
            return Task.FromResult(new OfficeAiExecutionResponse(Field("total", id == "e1" ? "42" : "45", id)));
        });
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(document, Request() with {
            Operation = OfficeAiOperation.ExtractFields, Fields = new[] { new OfficeAiFieldDefinition("total", OfficeAiFieldType.Decimal) }
        });
        OfficeAiField field = Assert.Single(result.Fields);
        Assert.Equal(2, executor.Requests.Count);
        Assert.Equal(OfficeAiFieldStatus.Conflicting, field.Status);
        Assert.Null(field.NormalizedValue);
        Assert.Null(field.RawValue);
        Assert.Equal(2, field.Citations.Count);
    }

    [Fact]
    public async Task ImageParseReopensThroughReaderAndKeepsProposalWarning() {
        byte[] bytes = { 1, 2, 3 };
        var image = new OfficeAiImage("image-page-1", 1, "image/png", bytes, 20, 20);
        bytes[0] = 9;
        byte[] copy = image.CopyBytes(); copy[1] = 9;
        OfficeAiDocument document = OfficeAiDocument.FromReadResult(new byte[] { 4 }, new(), new[] { image });
        const string parsed = """
            {"status":"ok","claims":[],"fields":[],"blocks":[{"kind":"heading","text":"Stock","evidence":[{"id":"image-page-1","quote":null}]}],
             "tables":[{"title":"Stock","columns":["Item","Count"],"rows":[["Pencil","12"]],"evidence":[{"id":"image-page-1","quote":null}]}]}
            """;
        var executor = new Executor(parsed, images: true);
        OfficeAiResult result = await new OfficeAiEngine(executor).RunAsync(document, Request() with { Operation = OfficeAiOperation.Parse, IncludeImages = true });
        Assert.Equal(new byte[] { 1, 2, 3 }, Assert.Single(Assert.Single(executor.Requests).Images).CopyBytes());
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        Assert.False(Assert.Single(Assert.Single(result.Tables).Citations).QuoteMatched);
        OfficeDocumentReadResult reopened = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(OfficeAiArtifacts.CreateProposedReadResult(document, result)));
        Assert.Equal(new[] { "Pencil", "12" }, Assert.Single(Assert.Single(reopened.Tables).Rows));
        Assert.Equal("ai-proposed-requires-review", Assert.Single(reopened.Diagnostics).Code);
        Assert.Equal(document.SourceHash, reopened.Source.SourceHash);
        Assert.Equal(1, Assert.Single(reopened.Blocks).Location.Page);
    }

    [Fact]
    public async Task StructuralParserRejectsRaggedTables() {
        const string parsed = """
            {"status":"ok","claims":[],"fields":[],"blocks":[],"tables":[{"title":"","columns":["A","B"],"rows":[["only one"]],"evidence":[{"id":"e1","quote":"table"}]}]}
            """;
        OfficeAiResult result = await new OfficeAiEngine(new Executor(parsed)).RunAsync(Document("table"), Request() with { Operation = OfficeAiOperation.Parse });
        Assert.Equal(OfficeAiResultStatus.InvalidResponse, result.Status);
        Assert.Empty(result.Tables);
    }

    [Fact]
    public async Task UnknownPageIsRejectedAndEmptyPageIsReported() {
        OfficeAiDocument document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, new OfficeDocumentReadResult { Pages = new[] { new OfficeDocumentPage { Number = 1 } } });
        var executor = new Executor(Empty);
        var engine = new OfficeAiEngine(executor);
        await Assert.ThrowsAsync<ArgumentException>(() => engine.RunAsync(document, Request() with { Pages = new[] { 2 } }));
        OfficeAiResult result = await engine.RunAsync(document, Request());
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Equal(new[] { 1 }, result.EmptyPages);
        Assert.Empty(executor.Requests);
    }

    [Fact]
    public void NestedPageLocationAndGeometryAreCapturedDefensively() {
        var region = new OfficeDocumentRegion { X = 5, Y = 10, Width = 20, Height = 30 };
        OfficeAiDocument document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, new OfficeDocumentReadResult {
            Pages = new[] { new OfficeDocumentPage { Number = 3, Blocks = new[] { new OfficeDocumentBlock { Text = "nested", Region = region } } } }
        });
        region.Width = 99;
        OfficeAiEvidence item = Assert.Single(document.Evidence);
        Assert.Equal(3, item.Page);
        Assert.Equal(20, item.Region!.Width);
    }

    [Fact]
    public async Task SmallerOperationBudgetRejectsPreviouslyCapturedLargerSource() {
        await Assert.ThrowsAsync<ArgumentException>(() => new OfficeAiEngine(new Executor(Empty)).RunAsync(Document("more than one byte"),
            Request() with { Limits = new OfficeAiLimits { MaxInputBytes = 1 } }));
    }
    [Fact]
    public async Task ArtifactsRejectDifferentEvidenceFromIdenticalSourceBytes() {
        byte[] bytes = Encoding.UTF8.GetBytes("Total 42. Other 99.");
        var first = OfficeAiDocument.FromReadResult(bytes, new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Text = "Total 42" } } });
        var second = OfficeAiDocument.FromReadResult(bytes, new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Text = "Other 99" } } });
        var result = await new OfficeAiEngine(new Executor(Claim("e1", "Total 42"))).RunAsync(first, Request());
        Assert.Equal(first.SourceHash, second.SourceHash);
        Assert.Throws<ArgumentException>(() => OfficeAiArtifacts.SerializeReport(second, result));
        Assert.Throws<ArgumentException>(() => OfficeAiArtifacts.CreateProposedReadResult(second, result with { Operation = OfficeAiOperation.Parse }));
    }

    [Fact]
    public void SnapshotFingerprintIncludesImagesAndIsStableForEquivalentEvidence() {
        byte[] bytes = new byte[] { 1 };
        var read = new OfficeDocumentReadResult();
        OfficeAiDocument Capture(byte pixel) => OfficeAiDocument.FromReadResult(bytes, read,
            new[] { new OfficeAiImage("page", 1, "image/png", new byte[] { pixel }, 1, 1) });
        Assert.Equal(Capture(1).SnapshotHash, Capture(1).SnapshotHash);
        Assert.NotEqual(Capture(1).SnapshotHash, Capture(2).SnapshotHash);
        var noImages = OfficeAiDocument.FromReadResult(bytes, read);
        Assert.NotEqual(Capture(1).SnapshotHash, noImages.SnapshotHash);
    }

    [Theory]
    [InlineData("chunk-warning")]
    [InlineData("chunk-table")]
    [InlineData("table-count")]
    public void AlternateReaderIncompletenessSignalsAreRetained(string signal) {
        var table = new ReaderTable { Columns = new[] { "Item" }, Rows = new[] { new[] { "retained" } } };
        var read = new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Text = "retained" } } };
        if (signal == "chunk-warning") read.Chunks = new[] { new ReaderChunk { Text = "retained", Warnings = new[] { "truncated" } } };
        else if (signal == "chunk-table") { table.Truncated = true; read.Chunks = new[] { new ReaderChunk { Tables = new[] { table } } }; }
        else {
            read.Tables = new[] { table };
            table.TotalRowCount = 2;
        }
        Assert.True(OfficeAiDocument.FromReadResult(new byte[] { 1 }, read).HasSourceDiagnostics);
    }

    [Fact]
    public void SourceGeometryRowCountCanIncludeAHeaderWithoutBodyTruncation() {
        var table = new ReaderTable {
            Columns = new[] { "Item" }, Rows = new[] { new[] { "first" }, new[] { "second" } }, TotalRowCount = 2,
            Diagnostics = new ReaderTableDiagnostics { SourceRowCount = 3 }
        };
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, new OfficeDocumentReadResult { Tables = new[] { table } });
        Assert.False(document.HasSourceDiagnostics);
        Assert.Equal(2, document.Evidence.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TruncatedTableWithoutTopLevelDiagnosticsPreservesIncompleteCoverage(bool pageOwned) {
        var table = new ReaderTable { Columns = new[] { "Item" }, Rows = new[] { new[] { "retained" } }, Truncated = true, TotalRowCount = 2 };
        var source = new OfficeDocumentReadResult();
        if (pageOwned) source.Pages = new[] { new OfficeDocumentPage { Number = 1, Tables = new[] { table } } };
        else source.Tables = new[] { table };
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, source);
        Assert.True(document.HasSourceDiagnostics);
        const string missing = """
            {"status":"insufficient","claims":[],"fields":[{"name":"value","status":"missing","rawValue":null,"evidence":[]}],"blocks":[],"tables":[]}
            """;
        var result = await new OfficeAiEngine(new Executor(missing)).RunAsync(document, Request() with {
            Operation = OfficeAiOperation.ExtractFields, Fields = new[] { new OfficeAiFieldDefinition("value") }
        });
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Equal(OfficeAiFieldStatus.NotEvaluated, Assert.Single(result.Fields).Status);
    }

    private static OfficeAiDocument Document(string text) => OfficeAiDocument.FromReadResult(Encoding.UTF8.GetBytes(text),
        new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Text = text, Kind = "paragraph", Location = new() { Page = 1 } } } });
    private static string Claim(string id, string? quote) => JsonSerializer.Serialize(new {
        status = "ok", claims = new[] { new { text = "The total is 42.", evidence = new[] { new { id, quote } } } },
        fields = Array.Empty<object>(), blocks = Array.Empty<object>(), tables = Array.Empty<object>()
    });
    private static string Field(string name, string rawValue, string id) => JsonSerializer.Serialize(new {
        status = "ok", claims = Array.Empty<object>(), fields = new[] { new { name, status = "present", rawValue, evidence = new[] { new { id, quote = rawValue } } } },
        blocks = Array.Empty<object>(), tables = Array.Empty<object>()
    });
    private sealed class Executor : IOfficeAiExecutor {
        private readonly Func<OfficeAiExecutionRequest, CancellationToken, Task<OfficeAiExecutionResponse>> _run;
        public Executor(string json, bool local = true, bool images = false) : this(new OfficeAiExecutionResponse(json), local, images) { }
        public Executor(OfficeAiExecutionResponse response, bool local = true, bool images = false) : this((_, _) => Task.FromResult(response), local, images) { }
        public Executor(Func<OfficeAiExecutionRequest, CancellationToken, Task<OfficeAiExecutionResponse>> run, bool local = true, bool images = false) {
            _run = run; Profile = new() { Id = "contract", Provider = "fixture", Model = "fixture", IsLocal = local, SupportsImages = images };
        }
        public OfficeAiExecutionProfile Profile { get; }
        public List<OfficeAiExecutionRequest> Requests { get; } = new();
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            Requests.Add(request); return _run(request, cancellationToken);
        }
    }
}
