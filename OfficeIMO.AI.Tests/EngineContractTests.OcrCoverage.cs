using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed partial class EngineContractTests {
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
