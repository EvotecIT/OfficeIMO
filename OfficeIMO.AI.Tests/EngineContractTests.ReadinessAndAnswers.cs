using System.Net;
using OfficeIMO.AI.IntelligenceX;
using OfficeIMO.Reader;
using OfficeIMO.Reader.DocBook;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed partial class EngineContractTests {
    [Theory]
    [InlineData(OfficeAiOperation.Ask)]
    [InlineData(OfficeAiOperation.Summarize)]
    public async Task AuthenticationFailureStopsRemainingEvidenceBatches(OfficeAiOperation operation) {
        var executor = new Executor((_, _) => throw new OfficeAiExecutionException(OfficeAiExecutionFailure.AuthenticationRequired));
        var result = await new OfficeAiEngine(executor).RunAsync(Document(new string('a', 100000)), Request() with { Operation = operation });
        Assert.Single(executor.Requests);
        Assert.Contains("provider-authentication-required", result.Diagnostics);
        Assert.NotEmpty(result.OmittedEvidenceIds);
        Assert.Empty(result.ProcessedEvidenceIds);
    }

    [Theory]
    [InlineData("warning")]
    [InlineData("visual")]
    public void GeneratedReaderNoticesCannotBecomeCitableEvidence(string kind) {
        var document = OfficeAiDocument.FromReadResult([1], new OfficeDocumentReadResult {
            Pages = [new() { Number = 1 }],
            Chunks = [new() { Kind = ReaderInputKind.Pdf, Text = "Generated reader notice", Location = new() { Page = 1, SourceBlockKind = kind } }]
        });
        Assert.Empty(document.Evidence);
        Assert.Equal([1], OfficeAiEvidenceReadiness.Inspect(document).PagesWithoutText);
    }

    [Fact]
    public async Task DocBookWarningRemainsCitableSourceContent() {
        const string xml = "<article xmlns=\"http://docbook.org/ns/docbook\" version=\"5.2\"><warning><para>Disconnect power before servicing.</para></warning></article>";
        var reader = new OfficeDocumentReaderBuilder().AddDocBookHandler().Build();
        var chunks = reader.ReadDocument(System.Text.Encoding.UTF8.GetBytes(xml), "safety.docbook").Chunks;
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "warning");
        var document = OfficeAiDocument.FromReadResult(System.Text.Encoding.UTF8.GetBytes(xml), new() { Chunks = chunks });
        var warning = Assert.Single(document.Evidence);
        Assert.Equal("Disconnect power before servicing.", warning.Text);
        var executor = new Executor("""{"claims":[{"text":"Disconnect the power first.","evidence":[{"id":"e1","quote":"Disconnect power before servicing."}]}],"fields":[],"blocks":[],"tables":[]}""");
        var result = await new OfficeAiEngine(executor).RunAsync(document, Request());
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        Assert.Equal(warning.Id, Assert.Single(Assert.Single(result.Claims).Citations).EvidenceId);
    }

    [Fact]
    public void TextReadinessKeepsEmptyPagesAndScopeVisible() {
        var document = OfficeAiDocument.FromReadResult([1], new OfficeDocumentReadResult {
            Pages = [new() { Number = 1 }, new() { Number = 2 }],
            Blocks = [new() { Text = "Total 42", Location = new() { Page = 1 } }]
        });
        var whole = OfficeAiEvidenceReadiness.Inspect(document);
        Assert.True(whole.HasText);
        Assert.Equal(8, whole.TextCharacters);
        Assert.Equal([2], whole.PagesWithoutText);
        var empty = OfficeAiEvidenceReadiness.Inspect(document, [2]);
        Assert.False(empty.HasText);
        Assert.Equal([2], empty.Pages);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeAiEvidenceReadiness.Inspect(document, [3]));
    }

    [Fact]
    public async Task AnswerArtifactPreservesCitationsAndRejectsAnotherSnapshot() {
        var document = Document("Total 42");
        var executor = new Executor("""{"claims":[{"text":"The total is 42.","evidence":[{"id":"e1","quote":"42"}]}],"fields":[],"blocks":[],"tables":[]}""");
        var result = await new OfficeAiEngine(executor).RunAsync(document, Request());
        string text = OfficeAiArtifacts.FormatAnswer(document, result, "What is the total?", "invoice.pdf");
        Assert.Contains(document.SourceHash, text);
        Assert.Contains(document.SnapshotHash, text);
        Assert.Contains("The total is 42.", text);
        Assert.Contains("Quote: 42", text);
        Assert.Contains("page 1", text);
        Assert.Contains("requires review", text);
        Assert.Throws<ArgumentException>(() => OfficeAiArtifacts.FormatAnswer(Document("Other 99"), result, "Question", "other.pdf"));
    }

    [Theory]
    [InlineData(401, OfficeAiExecutionFailure.AuthenticationRequired, "provider-authentication-required")]
    [InlineData(403, OfficeAiExecutionFailure.AccessDenied, "provider-access-denied")]
    [InlineData(404, OfficeAiExecutionFailure.NotFound, "provider-not-found")]
    [InlineData(429, OfficeAiExecutionFailure.RateLimited, "provider-rate-limited")]
    [InlineData(503, OfficeAiExecutionFailure.Unavailable, "provider-unavailable")]
    public async Task TypedProviderFailuresKeepActionableCodesWithoutProviderContent(int status, OfficeAiExecutionFailure expected, string code) {
        var raw = new HttpRequestException("secret-token and private provider payload", null, (HttpStatusCode)status);
        var classification = IntelligenceXOfficeAiErrors.Classify(raw);
        Assert.Equal(expected, classification);
        var safe = new OfficeAiExecutionException(classification);
        var executor = new Executor((_, _) => Task.FromException<OfficeAiExecutionResponse>(safe));
        var result = await new OfficeAiEngine(executor).RunAsync(Document("Total 42"), Request());
        Assert.Contains(code, result.Diagnostics);
        Assert.DoesNotContain("secret-token", OfficeAiArtifacts.SerializeReport(Document("Total 42"), result));
        Assert.Null(safe.InnerException);
        Assert.Equal(OfficeAiExecutionFailure.Unknown, IntelligenceXOfficeAiErrors.Classify(new InvalidOperationException("HTTP 401 secret-token")));
    }
}
