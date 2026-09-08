using System.Text;
using System.Text.Json;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed partial class EngineContractTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task InvalidUsageCannotProduceKnownTotals(bool providerThrows) {
        var executor = new Executor((_, _) => providerThrows
            ? throw new InvalidDataException("Invalid executor payload")
            : Task.FromResult(new OfficeAiExecutionResponse(Empty, InputTokens: -1, OutputTokens: 2)));
        var result = await new OfficeAiEngine(executor).RunAsync(Document("source"), Request());
        Assert.Equal(OfficeAiResultStatus.InvalidResponse, result.Status);
        Assert.Null(result.InputTokens);
        Assert.Null(result.OutputTokens);
    }

    [Theory]
    [InlineData(OfficeAiOperation.Ask)]
    [InlineData(OfficeAiOperation.ExtractFields)]
    [InlineData(OfficeAiOperation.Parse)]
    public async Task GenerationSchemaRequiresNonblankContentAndEvidence(OfficeAiOperation operation) {
        var executor = new Executor(Empty);
        await new OfficeAiEngine(executor).RunAsync(Document("source"), Request() with {
            Operation = operation, Fields = operation == OfficeAiOperation.ExtractFields
                ? new[] { new OfficeAiFieldDefinition("value") } : Array.Empty<OfficeAiFieldDefinition>()
        });
        using var schema = JsonDocument.Parse(Assert.Single(executor.Requests).OutputSchema);
        var properties = schema.RootElement.GetProperty("properties");
        var requiredText = operation == OfficeAiOperation.ExtractFields
            ? schema.RootElement.GetProperty("$defs").GetProperty("fieldValue").GetProperty("anyOf")[0].GetProperty("properties")
            : properties.GetProperty(operation == OfficeAiOperation.Parse ? "blocks" : "claims").GetProperty("items").GetProperty("properties");
        var stringSchemas = new[] { requiredText.GetProperty(operation == OfficeAiOperation.ExtractFields ? "rawValue" : "text"),
            requiredText.GetProperty("evidence").GetProperty("items").GetProperty("properties").GetProperty("id"),
            requiredText.GetProperty("evidence").GetProperty("items").GetProperty("properties").GetProperty("quote") };
        foreach (var textSchema in stringSchemas) {
            Assert.Equal(1, textSchema.GetProperty("minLength").GetInt32());
            string pattern = textSchema.GetProperty("pattern").GetString()!;
            foreach (string blank in new[] { "", " ", "\t\r\n", "\u0085\u00a0\u202f", "\u3000" })
                Assert.DoesNotMatch(pattern, blank);
            Assert.Matches(pattern, " source ");
            Assert.DoesNotMatch(pattern, "source\0text");
        }
        if (operation == OfficeAiOperation.Parse) {
            var cell = schema.RootElement.GetProperty("$defs").GetProperty("text");
            Assert.False(cell.TryGetProperty("minLength", out _));
            Assert.Matches(cell.GetProperty("pattern").GetString()!, "");
            Assert.DoesNotMatch(cell.GetProperty("pattern").GetString()!, "cell\0text");
        }
    }

    [Theory]
    [InlineData("Running")]
    [InlineData("Validating")]
    public async Task ProgressResourceFailureIsNotSwallowed(string stage) {
        var executor = new Executor(Empty);
        await Assert.ThrowsAsync<OutOfMemoryException>(() => new OfficeAiEngine(executor).RunAsync(Document("source"), Request(),
            progress: new FatalProgress(stage)));
    }

    private sealed class FatalProgress(string stage) : IProgress<OfficeAiProgress> {
        public void Report(OfficeAiProgress value) { if (value.Stage == stage) throw new OutOfMemoryException("fixture resource failure"); }
    }

    [Fact]
    public async Task NumericCitationSelectsTheCompleteOccurrenceAndRecordsItsOffset() {
        const string source = "1234 ... 234";
        var result = await new OfficeAiEngine(new Executor(Field("234", "e1"))).RunAsync(Document(source), Request() with {
            Operation = OfficeAiOperation.ExtractFields,
            Fields = new[] { new OfficeAiFieldDefinition("amount", OfficeAiFieldType.Integer) }
        });
        var field = Assert.Single(result.Fields);
        Assert.Equal(OfficeAiFieldStatus.Present, field.Status);
        Assert.Equal("234", field.NormalizedValue);
        Assert.Equal(source.LastIndexOf("234", StringComparison.Ordinal), Assert.Single(field.Citations).QuoteStart);
    }

    [Fact]
    public async Task ProviderOutOfMemoryStopsBeforeTheNextBatch() {
        var failure = new OutOfMemoryException("fixture resource failure");
        var executor = new Executor((_, _) => throw failure);
        Assert.Same(failure, await Assert.ThrowsAsync<OutOfMemoryException>(() =>
            new OfficeAiEngine(executor).RunAsync(Document(new string('a', 100000)), Request())));
        Assert.Single(executor.Requests);
    }

    [Theory]
    [InlineData(OfficeDocumentDiagnosticSeverity.Information, OfficeDocumentDiagnosticCategory.Detection, "input-kind-detected", false)]
    [InlineData(OfficeDocumentDiagnosticSeverity.Warning, OfficeDocumentDiagnosticCategory.Parsing, "partial-parse", true)]
    [InlineData(OfficeDocumentDiagnosticSeverity.Information, OfficeDocumentDiagnosticCategory.Content, "omitted-content", true)]
    public async Task SourceCoverageDistinguishesDetectionInformationFromOmissions(OfficeDocumentDiagnosticSeverity severity,
        OfficeDocumentDiagnosticCategory category, string code, bool incomplete) {
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, new OfficeDocumentReadResult {
            Blocks = new[] { new OfficeDocumentBlock { Text = "Total 42" } },
            Diagnostics = new[] { new OfficeDocumentDiagnostic { Severity = severity, Category = category, Code = code } }
        });
        var result = await new OfficeAiEngine(new Executor(Claim("e1", "Total 42"))).RunAsync(document, Request());
        Assert.Equal(incomplete, document.HasSourceDiagnostics);
        Assert.Equal(incomplete ? OfficeAiResultStatus.Partial : OfficeAiResultStatus.Completed, result.Status);
    }

    [Theory]
    [InlineData(32000, false, true)]
    [InlineData(32001, false, false)]
    [InlineData(32000, true, true)]
    public async Task GeneratedStringLimitMatchesLocalUnicodeValidation(int length, bool supplementary, bool valid) {
        string text = supplementary ? string.Concat(Enumerable.Repeat("😀", length)) : new string('x', length);
        string response = JsonSerializer.Serialize(new {
            status = "ok", claims = new[] { new { text, evidence = new[] { new { id = "e1", quote = "source" } } } },
            fields = Array.Empty<object>(), blocks = Array.Empty<object>(), tables = Array.Empty<object>()
        });
        var executor = new Executor(response);
        var result = await new OfficeAiEngine(executor).RunAsync(Document("source"), Request() with { Limits = new() { MaxResponseCharacters = 500000 } });
        using var schema = JsonDocument.Parse(Assert.Single(executor.Requests).OutputSchema);
        var textSchema = schema.RootElement.GetProperty("properties").GetProperty("claims").GetProperty("items").GetProperty("properties").GetProperty("text");
        int maximum = textSchema.GetProperty("maxLength").GetInt32();
        Assert.Equal(valid, text.EnumerateRunes().Count() <= maximum);
        Assert.Equal(valid ? OfficeAiResultStatus.Completed : OfficeAiResultStatus.InvalidResponse, result.Status);
    }

    [Theory]
    [InlineData("Running")]
    [InlineData("Validating")]
    [InlineData("Completed")]
    public async Task ThrowingProgressObserversCannotReplaceAValidOperation(string failingStage) {
        var executor = new Executor(Claim("e1", "Total 42"));
        var result = await new OfficeAiEngine(executor).RunAsync(Document("Total 42"), Request(), new ThrowingProgress(failingStage));
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        Assert.Single(result.Claims);
        Assert.DoesNotContain("provider-execution-failed", result.Diagnostics);
        Assert.Single(executor.Requests);
    }

    private sealed class ThrowingProgress(string stage) : IProgress<OfficeAiProgress> {
        public void Report(OfficeAiProgress value) {
            if (value.Stage == stage) throw new InvalidOperationException("Observer failed");
        }
    }

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
            Limits = new() { MaxResultItems = 2 },
            Fields = operation == OfficeAiOperation.ExtractFields ? new[] { new OfficeAiFieldDefinition("total") } : Array.Empty<OfficeAiFieldDefinition>()
        });
        using JsonDocument schema = JsonDocument.Parse(Assert.Single(executor.Requests).OutputSchema);
        string[] enabled = enabledNames.Split(',');
        foreach (string name in new[] { "claims", "fields", "blocks", "tables" }) {
            JsonElement array = schema.RootElement.GetProperty("properties").GetProperty(name);
            if (name == "fields" && operation == OfficeAiOperation.ExtractFields) {
                Assert.Equal("object", array.GetProperty("type").GetString());
                Assert.False(array.GetProperty("additionalProperties").GetBoolean());
                Assert.Equal("field1", Assert.Single(array.GetProperty("required").EnumerateArray()).GetString());
                Assert.Equal("field1", Assert.Single(array.GetProperty("properties").EnumerateObject()).Name);
                continue;
            }
            Assert.Equal("array", array.GetProperty("type").GetString());
            Assert.Equal(enabled.Contains(name) ? 2 : 0, array.GetProperty("maxItems").GetInt32());
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
        var executor = new Executor(Field(raw, "e1"));
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

    [Theory]
    [InlineData("Refund: -1 234,50 PLN.", "1 234,50", "pl-PL", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: - 42 PLN.", "42", "pl-PL", OfficeAiFieldType.Integer)]
    [InlineData("Total: 142.", "42", "en-US", OfficeAiFieldType.Integer)]
    [InlineData("Total: 42.50 USD.", "42", "en-US", OfficeAiFieldType.Integer)]
    [InlineData("Total: 1,234.50 USD.", "234.50", "en-US", OfficeAiFieldType.Decimal)]
    [InlineData("Total: 1 234,50 PLN.", "234,50", "pl-PL", OfficeAiFieldType.Decimal)]
    [InlineData("Total: .50 USD.", "50", "en-US", OfficeAiFieldType.Decimal)]
    [InlineData("Total: 42- USD.", "42", "en-US", OfficeAiFieldType.Integer)]
    [InlineData("Total: 42e3 USD.", "42", "en-US", OfficeAiFieldType.Integer)]
    [InlineData("Total: 42E\u22123 USD.", "42", "en-US", OfficeAiFieldType.Integer)]
    [InlineData("Total: 42e\u22123 USD.", "42", "en-US", OfficeAiFieldType.Decimal)]
    [InlineData("Total: 42E\u061c-3.", "42", "ar-EG", OfficeAiFieldType.Integer)]
    [InlineData("Total: 42E\u061c+3.", "42", "ar-EG", OfficeAiFieldType.Integer)]
    [InlineData("Total: 42e3 USD.", "3", "en-US", OfficeAiFieldType.Integer)]
    [InlineData("Total: 42e-3 USD.", "-3", "en-US", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: -$42.00", "42.00", "en-US", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: -usd 42.00", "42.00", "en-US", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: 42.00 uSd-", "42.00", "en-US", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: -ZŁ 42,00", "42,00", "pl-PL", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: -R$42,00", "42,00", "pt-BR", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: 42,00R$-", "42,00", "pt-BR", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: -R$42.00", "42.00", "en-US", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: -US$42.00", "42.00", "en-US", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: 42.00 $-", "42.00", "en-US", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: ($42.00)", "42.00", "en-US", OfficeAiFieldType.Decimal)]
    [InlineData("Total: ,50 PLN.", "50", "pl-PL", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: -PLN 42,00", "42,00", "pl-PL", OfficeAiFieldType.Decimal)]
    [InlineData("Refund: 42,00 zł-", "42,00", "pl-PL", OfficeAiFieldType.Decimal)]
    public async Task NumericSubstringCannotBecomeAnAcceptedAmount(string source, string raw, string culture, OfficeAiFieldType type) {
        foreach (string quote in new[] { raw, source }) {
            string response = JsonSerializer.Serialize(new {
                status = "ok", claims = Array.Empty<object>(),
                fields = new { field1 = new { status = "present", rawValue = raw, evidence = new[] { new { id = "e1", quote } } } },
                blocks = Array.Empty<object>(), tables = Array.Empty<object>()
            });
            var result = await new OfficeAiEngine(new Executor(response)).RunAsync(Document(source), Request() with {
                Operation = OfficeAiOperation.ExtractFields, Culture = culture,
                Fields = new[] { new OfficeAiFieldDefinition("amount", type) }
            });
            Assert.Equal(OfficeAiFieldStatus.Invalid, Assert.Single(result.Fields).Status);
            Assert.Null(Assert.Single(result.Fields).NormalizedValue);
        }
    }

    [Theory]
    [InlineData("Refund: -1 234,50 PLN.", "-1 234,50", "pl-PL", "-1234.50")]
    [InlineData("Total: 42.50USD.", "42.50", "en-US", "42.50")]
    [InlineData("Total: 42. Next item.", "42", "en-US", "42")]
    [InlineData("Total: .50 USD.", ".50", "en-US", "0.50")]
    [InlineData("Total: 42 pre-tax", "42", "en-US", "42")]
    [InlineData("Count: 42 red-green widgets", "42", "en-US", "42")]
    [InlineData("Count: 42 top-rated widgets", "42", "en-US", "42")]
    [InlineData("Count: 42 all-in", "42", "en-US", "42")]
    [InlineData("Count: 42 usd per item", "42", "en-US", "42")]
    [InlineData("Quantity: 42 2025 model", "42", "en-US", "42")]
    [InlineData("Quantity: 42\u00a02025 model", "42", "en-US", "42")]
    [InlineData("Quantity: 42\u202f2025 model", "42", "en-US", "42")]
    [InlineData("Quantity: 2025 42 units", "42", "en-US", "42")]
    public async Task CompleteNumericEvidencePreservesSignsUnitsAndSentencePunctuation(string source, string raw, string culture, string expected) {
        var result = await new OfficeAiEngine(new Executor(Field(raw, "e1"))).RunAsync(Document(source), Request() with {
            Operation = OfficeAiOperation.ExtractFields, Culture = culture,
            Fields = new[] { new OfficeAiFieldDefinition("amount", OfficeAiFieldType.Decimal) }
        });
        Assert.Equal(OfficeAiFieldStatus.Present, Assert.Single(result.Fields).Status);
        Assert.Equal(expected, Assert.Single(result.Fields).NormalizedValue);
    }

    private static OfficeAiRequest Request() => new() { Instruction = "What is the total?" };
    [Theory]
    [InlineData("1,234", "en-US", "1234")]
    [InlineData("-1,234", "en-US", "-1234")]
    [InlineData("1 234", "pl-PL", "1234")]
    [InlineData("12,34,567", "hi-IN", "1234567")]
    [InlineData("1,23", "en-US", null)]
    [InlineData("1,,234", "en-US", null)]
    [InlineData("123,456", "hi-IN", null)]
    [InlineData("1,234.0", "en-US", null)]
    [InlineData("9,223,372,036,854,775,808", "en-US", null)]
    public async Task IntegerFieldsValidateLocaleGroupingAndIntegralRange(string raw, string culture, string? expected) {
        var result = await new OfficeAiEngine(new Executor(Field(raw, "e1"))).RunAsync(Document(raw), Request() with {
            Operation = OfficeAiOperation.ExtractFields, Culture = culture,
            Fields = new[] { new OfficeAiFieldDefinition("amount", OfficeAiFieldType.Integer) }
        });
        var field = Assert.Single(result.Fields);
        Assert.Equal(expected is null ? OfficeAiFieldStatus.Invalid : OfficeAiFieldStatus.Present, field.Status);
        Assert.Equal(expected, field.NormalizedValue);
    }
    [Theory]
    [InlineData("1,23")]
    [InlineData("1,,234")]
    public async Task MalformedLocaleGroupingDoesNotSilentlyChangeTheAmount(string raw) {
        OfficeAiResult result = await new OfficeAiEngine(new Executor(Field(raw, "e1"))).RunAsync(Document(raw), Request() with {
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
            return Task.FromResult(new OfficeAiExecutionResponse(Field(id == "e1" ? "42" : "45", id)));
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
        Assert.Single(document.Evidence, item => item.Kind == "table");
        Assert.Equal(2, document.Evidence.Count(item => item.Kind == "table-row"));
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
            {"status":"insufficient","claims":[],"fields":{"field1":{"status":"missing","rawValue":null,"evidence":[]}},"blocks":[],"tables":[]}
            """;
        var result = await new OfficeAiEngine(new Executor(missing)).RunAsync(document, Request() with {
            Operation = OfficeAiOperation.ExtractFields, Fields = new[] { new OfficeAiFieldDefinition("value") }
        });
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Equal(OfficeAiFieldStatus.NotEvaluated, Assert.Single(result.Fields).Status);
    }

    [Theory]
    [InlineData("yyyy-MM-dd'")]
    [InlineData("yyyy-MM-dd\\")]
    [InlineData("q")]
    public async Task InvalidDateFormatIsRejectedBeforeInference(string format) {
        var executor = new Executor(Empty);
        await Assert.ThrowsAsync<ArgumentException>(() => new OfficeAiEngine(executor).RunAsync(Document("2030-04-03"), Request() with {
            Operation = OfficeAiOperation.ExtractFields,
            Fields = new[] { new OfficeAiFieldDefinition("date", OfficeAiFieldType.Date, format) }
        }));
        Assert.Empty(executor.Requests);
    }

    [Theory]
    [InlineData("en-US", "MM/dd/yyyy", "04/03/2030")]
    [InlineData("pl-PL", "dd MMMM yyyy", "03 kwietnia 2030")]
    public async Task ValidCultureSpecificDateFormatStillNormalizes(string culture, string format, string raw) {
        var result = await new OfficeAiEngine(new Executor(Field(raw, "e1"))).RunAsync(Document(raw), Request() with {
            Operation = OfficeAiOperation.ExtractFields, Culture = culture,
            Fields = new[] { new OfficeAiFieldDefinition("date", OfficeAiFieldType.Date, format) }
        });
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        Assert.Equal("2030-04-03", Assert.Single(result.Fields).NormalizedValue);
    }

    [Theory]
    [InlineData(1, 3, 3, true)]
    [InlineData(3, 1, 1, true)]
    [InlineData(2, 2, 2, false)]
    [InlineData(4, 1, 1, false)]
    [InlineData(1, 3, 2, false)]
    [InlineData(1, 2, 3, false)]
    [InlineData(1, 2, 0, false)]
    [InlineData(0, 3, 3, true)]
    [InlineData(1, 4, 4, false)]
    public async Task GeneratedTableShapesAndValidatorApplyTheSameEffectiveBounds(int rowCount, int columnCount, int rowWidth, bool accepted) {
        string[] columns = Enumerable.Range(0, columnCount).Select(index => "Column " + index).ToArray();
        string[][] values = Enumerable.Range(0, rowCount).Select(_ => Enumerable.Repeat("value", rowWidth).ToArray()).ToArray();
        string response = JsonSerializer.Serialize(new { status = "ok", claims = Array.Empty<object>(), fields = Array.Empty<object>(),
            blocks = Array.Empty<object>(), tables = new[] { new { title = "", columns, rows = values, evidence = new[] { new { id = "e1", quote = "table" } } } } });
        var executor = new Executor(response);
        var result = await new OfficeAiEngine(executor).RunAsync(Document("table"), Request() with {
            Operation = OfficeAiOperation.Parse, Limits = new() { MaxResultItems = 3, MaxTableCells = 3, MaxTableColumns = 3 }
        });
        var sent = Assert.Single(executor.Requests);
        using var schema = JsonDocument.Parse(sent.OutputSchema);
        var properties = schema.RootElement.GetProperty("properties");
        Assert.Equal(3, properties.GetProperty("tables").GetProperty("maxItems").GetInt32());
        var definitions = schema.RootElement.GetProperty("$defs");
        JsonElement Resolve(JsonElement node) => node.TryGetProperty("$ref", out var reference)
            ? definitions.GetProperty(reference.GetString()!.Split('/')[^1]) : node;
        bool CountFits(JsonElement array, int count) => (!array.TryGetProperty("minItems", out var minimum) || count >= minimum.GetInt32())
            && (!array.TryGetProperty("maxItems", out var maximum) || count <= maximum.GetInt32());
        bool schemaAcceptsShape = properties.GetProperty("tables").GetProperty("items").GetProperty("anyOf").EnumerateArray().Any(shape => {
            var table = shape.GetProperty("properties");
            var rows = table.GetProperty("rows");
            return CountFits(Resolve(table.GetProperty("columns")), columnCount) && CountFits(rows, rowCount)
                && values.All(row => CountFits(Resolve(rows.GetProperty("items")), row.Length));
        });
        Assert.Equal(accepted, schemaAcceptsShape);
        Assert.Equal(accepted ? OfficeAiResultStatus.Completed : OfficeAiResultStatus.InvalidResponse, result.Status);
        using var input = JsonDocument.Parse(sent.InputJson);
        Assert.Equal(3, input.RootElement.GetProperty("resultLimits").GetProperty("maxTableCells").GetInt32());
    }

    private static OfficeAiDocument Document(string text) => OfficeAiDocument.FromReadResult(Encoding.UTF8.GetBytes(text),
        new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Text = text, Kind = "paragraph", Location = new() { Page = 1 } } } });
    private static string Claim(string id, string? quote) => JsonSerializer.Serialize(new {
        status = "ok", claims = new[] { new { text = "The total is 42.", evidence = new[] { new { id, quote } } } },
        fields = Array.Empty<object>(), blocks = Array.Empty<object>(), tables = Array.Empty<object>()
    });
    private static string Field(string rawValue, string id) => JsonSerializer.Serialize(new {
        status = "ok", claims = Array.Empty<object>(), fields = new Dictionary<string, object> { ["field1"] = new { status = "present", rawValue, evidence = new[] { new { id, quote = rawValue } } } },
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
