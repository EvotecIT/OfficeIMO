using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed partial class EngineContractTests {
    [Theory]
    [InlineData("ocr", true)]
    [InlineData("page-ocr", true)]
    [InlineData("table", true)]
    [InlineData("chunk-warning", true)]
    [InlineData("warning", true)]
    [InlineData("error", true)]
    [InlineData("detection", false)]
    public async Task NestedSourceLimitationsSurviveTransportWithoutDuplicatingFlattenedEvidence(string limitation, bool incomplete) {
        var child = new OfficeDocumentReadResult();
        switch (limitation) {
            case "ocr": child.OcrCandidates = new[] { new OfficeDocumentOcrCandidate { Id = "scan" } }; break;
            case "page-ocr": child.Pages = new[] { new OfficeDocumentPage { Number = 1,
                OcrCandidates = new[] { new OfficeDocumentOcrCandidate { Id = "scan" } } } }; break;
            case "table": child.Tables = new[] { new ReaderTable { Columns = new[] { "Total" }, TotalRowCount = 1 } }; break;
            case "chunk-warning": child.Chunks = new[] { new ReaderChunk { Warnings = new[] { "Child content omitted" } } }; break;
            default: child.Diagnostics = new[] { new OfficeDocumentDiagnostic {
                Code = "child-" + limitation,
                Severity = limitation == "error" ? OfficeDocumentDiagnosticSeverity.Error
                    : limitation == "detection" ? OfficeDocumentDiagnosticSeverity.Information : OfficeDocumentDiagnosticSeverity.Warning,
                Category = limitation == "detection" ? OfficeDocumentDiagnosticCategory.Detection : OfficeDocumentDiagnosticCategory.Parsing
            } }; break;
        }
        var source = new OfficeDocumentReadResult {
            Blocks = new[] { new OfficeDocumentBlock { Text = "readable" } },
            NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "attachment.zip", Document = new OfficeDocumentReadResult {
                NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "scan.pdf", Document = child } }
            } } }
        };
        var restored = OfficeDocumentReadResultJson.Deserialize(source.ToJson());
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, restored);

        Assert.Equal(incomplete, document.HasSourceDiagnostics);
        Assert.Equal("readable", Assert.Single(document.Evidence).Text);
        var result = await new OfficeAiEngine(new Executor(Claim("e1", "readable"))).RunAsync(document, Request());
        Assert.Equal(incomplete ? OfficeAiResultStatus.Partial : OfficeAiResultStatus.Completed, result.Status);
        restored.NestedDocuments = Array.Empty<OfficeDocumentNestedResult>();
        Assert.Equal(incomplete, document.HasSourceDiagnostics);
        var complete = OfficeAiDocument.FromReadResult(new byte[] { 1 }, restored);
        Assert.False(complete.HasSourceDiagnostics);
        Assert.Equal(incomplete, complete.SnapshotHash != document.SnapshotHash);
    }

    [Fact]
    public async Task PendingNestedOcrPreventsMissingFieldFromBecomingAClaimOfAbsence() {
        var source = new OfficeDocumentReadResult {
            Blocks = new[] { new OfficeDocumentBlock { Text = "readable" } },
            NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "scan.pdf", Document = new OfficeDocumentReadResult {
                OcrCandidates = new[] { new OfficeDocumentOcrCandidate { Id = "scan" } }
            } } }
        };
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, source);
        string response = System.Text.Json.JsonSerializer.Serialize(new {
            claims = Array.Empty<object>(),
            fields = new { field1 = new { status = "missing", rawValue = (string?)null, evidence = Array.Empty<object>() } },
            blocks = Array.Empty<object>(), tables = Array.Empty<object>()
        });
        var result = await new OfficeAiEngine(new Executor(response)).RunAsync(document, Request() with {
            Operation = OfficeAiOperation.ExtractFields,
            Fields = new[] { new OfficeAiFieldDefinition("amount", OfficeAiFieldType.Integer) }
        });
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Equal(OfficeAiFieldStatus.NotEvaluated, Assert.Single(result.Fields).Status);
        Assert.Contains("source-reader-diagnostics", result.Diagnostics);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task PendingOcrWithoutWarningsRetainsIncompleteCoverage(bool pageOwned, bool roundTrip) {
        var candidate = new OfficeDocumentOcrCandidate { Id = "ocr", Kind = "image", Location = new() { Page = 1 } };
        var source = new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Text = "readable", Location = new() { Page = 1 } } } };
        if (pageOwned) source.Pages = new[] { new OfficeDocumentPage { Number = 1, OcrCandidates = new[] { candidate } } };
        else source.OcrCandidates = new[] { candidate };
        var reader = roundTrip ? OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source)) : source;
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, reader);
        Assert.True(document.HasSourceDiagnostics);
        var result = await new OfficeAiEngine(new Executor(Claim("e1", "readable"))).RunAsync(document, Request());
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        source.OcrCandidates = Array.Empty<OfficeDocumentOcrCandidate>();
        source.Pages = Array.Empty<OfficeDocumentPage>();
        var complete = OfficeAiDocument.FromReadResult(new byte[] { 1 }, source);
        Assert.False(complete.HasSourceDiagnostics);
        Assert.NotEqual(document.SnapshotHash, complete.SnapshotHash);
    }
}
