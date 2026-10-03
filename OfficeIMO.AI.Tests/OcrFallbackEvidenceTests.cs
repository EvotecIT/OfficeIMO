using OfficeIMO.Ocr;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class OcrFallbackEvidenceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NativeFallbackEvidenceAndPdfFilteringSurviveDirectAndNestedOcr(bool tree) {
        var source = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Email, Source = new() { Path = "mail.eml" },
            Chunks = [
                new() { Id = "native", Text = "Native amount 7", Kind = ReaderInputKind.Email, Location = new() { Path = "mail.eml" } },
                new() { Id = "notice", Text = "Generated reader notice", Kind = ReaderInputKind.Pdf, Location = new() { Path = "table.pdf", Page = 1, SourceBlockKind = "warning" } },
                new() { Id = "table", Text = "A-100", Kind = ReaderInputKind.Pdf, Location = new() { Path = "table.pdf", Page = 1, SourceBlockKind = "table" }, Diagnostics = new() { TableCount = 1 } },
                new() { Id = "partial", Text = "Unprojected second table B-200", Kind = ReaderInputKind.Pdf, Location = new() { Path = "table.pdf", Page = 2, SourceBlockKind = "table" }, Diagnostics = new() { TableCount = 2 } }
            ],
            Tables = [
                new() { Columns = ["Code"], Rows = [["A-100"]], Location = new() { Path = "table.pdf", Page = 1 } },
                new() { Columns = ["Code"], Rows = [["C-300"]], Location = new() { Path = "table.pdf", Page = 2 } }
            ]
        };
        var candidateOwner = tree ? new OfficeDocumentReadResult { Source = new() { Path = "scan.png" } } : source;
        candidateOwner.OcrCandidates = [new() { Id = "scan", AssetId = "asset" }];
        candidateOwner.Assets = [new() { Id = "asset", Kind = "image", MediaType = "image/png", PayloadBytes = [1] }];
        if (tree) source.NestedDocuments = [new() { Path = "scan.png", Document = candidateOwner }];
        string before = source.ToJson();
        var engine = new DelegateOcrEngine("fallback", (_, _) => Task.FromResult(new OcrResult { Text = "Scanned amount 42" }));
        var enriched = tree ? await source.ApplyOcrTreeAsync(engine) : await source.ApplyOcrAsync(engine);
        var snapshot = OfficeAiDocument.FromReadResult([1], OfficeDocumentReadResultJson.Deserialize(enriched.Document.ToJson()));

        var native = Assert.Single(snapshot.Evidence, item => item.Text == "Native amount 7");
        Assert.Equal("mail.eml", native.SourceLocation!.Path);
        Assert.Single(snapshot.Evidence, item => item.Text == "Scanned amount 42");
        Assert.Single(snapshot.Evidence, item => item.Text.Contains("A-100"));
        Assert.Contains(snapshot.Evidence, item => item.Text.Contains("B-200"));
        Assert.DoesNotContain(snapshot.Evidence, item => item.Text == "Generated reader notice");
        Assert.Equal(before, source.ToJson());
    }
}
