using System.Text.Json;
using OfficeIMO.Ocr;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class RecognitionEvidenceTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task RecognitionSurvivesNestedTransportSnapshotAndExtraction(bool nested, bool disagree) {
        var scan = new OfficeDocumentReadResult {
            Source = new() { Path = "invoice.png" },
            OcrCandidates = [new() { Id = "scan", AssetId = "image", Location = new() { Page = 1 } }],
            Assets = [new() { Id = "image", Kind = "image", MediaType = "image/png", PayloadBytes = [1], Location = new() { Page = 1 } }]
        };
        var source = nested ? new OfficeDocumentReadResult { Source = new() { Path = "bundle.zip" },
            NestedDocuments = [new() { Path = "invoice.png", Document = scan }] } : scan;
        var engine = new AdaptiveOcrEngine("comparison", [Attempt("first", "42"), Attempt("second", disagree ? "99" : "42")],
            new OcrReviewPolicy(OcrRetryMode.CompareAll, minimumWordConfidence: .98));
        var execution = nested ? await source.ApplyOcrTreeAsync(engine) : await source.ApplyOcrAsync(engine);
        var transport = OfficeDocumentReadResultJson.Deserialize(execution.Document.ToJson());
        var document = OfficeAiDocument.FromReadResult([1], transport);
        var evidence = Assert.Single(document.Evidence);
        Assert.NotNull(evidence.Recognition);
        Assert.Equal("test-ocr", evidence.Recognition!.Provider);
        Assert.Equal("model-v1", evidence.Recognition.Model);
        Assert.Equal(2, evidence.Recognition.CompletedAttempts);
        Assert.True(evidence.Recognition.ConfidenceChecksPassed);
        Assert.Equal(.98, evidence.Recognition.MinimumWordConfidence);
        Assert.Equal(disagree, evidence.Recognition.HasDisagreement);
        Assert.Equal(disagree, evidence.Recognition.ReviewRecommended);
        Assert.Equal("invoice.png", Path.GetFileName(evidence.SourceLocation!.Path));
        var executor = new ExtractionExecutor();
        var result = await new OfficeAiEngine(executor).RunAsync(document, new() {
            Operation = OfficeAiOperation.ExtractFields, Instruction = "Extract the invoice total.",
            Fields = [new("total", OfficeAiFieldType.Decimal)]
        });
        var field = Assert.Single(result.Fields);
        Assert.Equal("42", field.NormalizedValue);
        Assert.True(field.TextValueMatched);
        Assert.Equal(disagree, field.RecognitionReviewRequired);
        Assert.Same(evidence.Recognition, Assert.Single(field.Citations).Recognition);
        using var request = JsonDocument.Parse(executor.Input!);
        Assert.Equal("test-ocr", request.RootElement.GetProperty("evidence")[0].GetProperty("recognition").GetProperty("Provider").GetString());
        using var report = JsonDocument.Parse(OfficeAiArtifacts.SerializeReport(document, result));
        Assert.Equal(disagree, report.RootElement.GetProperty("result").GetProperty("fields")[0].GetProperty("recognitionReviewRequired").GetBoolean());
        string originalHash = document.SnapshotHash;
        transport.Blocks[0].Recognition = new OfficeDocumentRecognitionEvidence(provider: "other-provider");
        Assert.Equal("test-ocr", evidence.Recognition.Provider);
        Assert.NotEqual(originalHash, OfficeAiDocument.FromReadResult([1], transport).SnapshotHash);
    }

    [Fact]
    public async Task DownstreamTruncationInvalidatesRecognitionChecks() {
        var source = new OfficeDocumentReadResult {
            OcrCandidates = [new() { Id = "scan", AssetId = "image" }],
            Assets = [new() { Id = "image", Kind = "image", MediaType = "image/png", PayloadBytes = [1] }]
        };
        var engine = new AdaptiveOcrEngine("comparison", [Attempt("first", "42"), Attempt("second", "42")],
            new OcrReviewPolicy(OcrRetryMode.CompareAll, minimumWordConfidence: .98));
        var execution = await source.ApplyOcrAsync(engine, new() { MaxRecognizedCharactersPerCandidate = 4 });
        var recognition = Assert.Single(execution.Document.Blocks).Recognition!;
        Assert.True(recognition.ReviewRecommended);
        Assert.False(recognition.ConfidenceChecksPassed);
        Assert.True(recognition.ComparisonIncomplete);
    }

    [Fact]
    public void SuppliedOcrWithoutAssessmentRemainsExplicitlyUnknownAfterTransport() {
        var source = new OfficeDocumentReadResult { OcrCandidates = [new() { Id = "scan" }] };
        var enriched = source.ApplyOcrResults([new() { CandidateId = "scan", Text = "Total 42", Provider = "external", Confidence = .99 }]).Document;
        var document = OfficeAiDocument.FromReadResult([1], OfficeDocumentReadResultJson.Deserialize(enriched.ToJson()));
        var recognition = Assert.Single(document.Evidence).Recognition!;
        Assert.Equal(.99, recognition.Confidence);
        Assert.Null(recognition.ReviewRecommended);
        Assert.Null(recognition.ConfidenceChecksPassed);
    }

    private static OcrRecognitionAttempt Attempt(string name, string value) => new(name,
        new DelegateOcrEngine("test-ocr", (_, _) => Task.FromResult(new OcrResult {
            Text = "Total " + value, Provider = "test-ocr", Model = "model-v1", Language = "eng", Confidence = .99,
            Spans = [new() { Text = "Total", Level = OcrTextSpanLevel.Word, Confidence = .99 },
                new() { Text = value, Level = OcrTextSpanLevel.Word, Confidence = .99 }]
        })));

    private sealed class ExtractionExecutor : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile { get; } = new() { Id = "test", Provider = "fixture", Model = "fixture", IsLocal = true };
        public string? Input { get; private set; }
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            Input = request.InputJson;
            using var input = JsonDocument.Parse(request.InputJson);
            string id = input.RootElement.GetProperty("evidence")[0].GetProperty("id").GetString()!;
            return Task.FromResult(new OfficeAiExecutionResponse(JsonSerializer.Serialize(new {
                claims = Array.Empty<object>(), blocks = Array.Empty<object>(), tables = Array.Empty<object>(),
                fields = new { field1 = new { status = "present", rawValue = "42", evidence = new[] { new { id, quote = "Total 42" } } } }
            })));
        }
    }
}
