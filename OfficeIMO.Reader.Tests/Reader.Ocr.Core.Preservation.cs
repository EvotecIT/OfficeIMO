using OfficeIMO.Ocr;
using OfficeIMO.Reader;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderOcrCoreTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ApplyOcrAsync_PreservesRichNestedDocumentsWithAndWithoutCandidates(bool recognize) {
        var child = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Source = new OfficeDocumentSource { Path = "attachment.pdf" },
            Assets = new[] { new OfficeDocumentAsset { Id = "image", Kind = "image", PayloadBytes = new byte[] { 1, 2 } } },
            Links = new[] { new OfficeDocumentLink { Id = "link", Kind = "uri", Uri = "https://example.test" } },
            Forms = new[] { new OfficeDocumentFormField { Id = "field", Kind = "text", Value = "retained" } },
            Metadata = new[] { new OfficeDocumentMetadataEntry { Id = "title", Value = "Attachment" } },
            Diagnostics = new[] { new OfficeDocumentDiagnostic { Code = "child-warning", Severity = OfficeDocumentDiagnosticSeverity.Warning } },
            NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "scan.png", Document = new OfficeDocumentReadResult {
                OcrCandidates = new[] { new OfficeDocumentOcrCandidate { Id = "pending" } }
            } } }
        };
        string childBefore = child.ToJson();
        var source = CreateDocument(recognize ? 1 : 0);
        source.NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "attachment.pdf", Document = child } };
        var engine = new DelegateOcrEngine("preservation", (_, _) => Task.FromResult(new OcrResult { Text = "Recognized" }));

        var execution = await source.ApplyOcrAsync(engine);

        Assert.Equal(recognize ? 1 : 0, execution.Report.RecognizedCandidateCount);
        Assert.Same(child, Assert.Single(execution.Document.NestedDocuments).Document);
        var restored = OfficeDocumentReadResultJson.Deserialize(execution.Document.ToJson());
        var nested = Assert.Single(restored.NestedDocuments);
        Assert.Equal("attachment.pdf", nested.Path);
        Assert.Equal(childBefore, nested.Document.ToJson());
        Assert.Equal(childBefore, child.ToJson());
        Assert.Equal(recognize ? 1 : 0, source.OcrCandidates.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ApplyOcrResults_PreservesLiteralHeadingsAndInvalidatesStaleHierarchyPaths(bool changeDisplayPath) {
        var location = new ReaderLocation { HeadingPath = "A > B", Page = 1 };
        ReaderHeadingPath.SetHierarchyPath(location, ReaderHeadingPath.Combine(new[] { "A > B" }));
        if (changeDisplayPath) location.HeadingPath = "C > D";
        var source = new OfficeDocumentReadResult {
            OcrCandidates = new[] { new OfficeDocumentOcrCandidate { Id = "scan", Location = location } }
        };

        var enriched = source.ApplyOcrResults(new[] { new OfficeDocumentOcrTextResult { CandidateId = "scan", Text = "Recognized" } }).Document;

        Assert.Equal(location.HierarchyHeadingPath, Assert.Single(enriched.Blocks).Location.HierarchyHeadingPath);
        var hierarchy = ReaderHierarchicalChunker.Chunk(enriched, new ReaderHierarchicalChunkingOptions { IncludeContextInText = false });
        Assert.Equal(changeDisplayPath ? new[] { "C", "D" } : new[] { "A > B" },
            hierarchy.Nodes.Where(node => node.Kind == ReaderChunkHierarchyNodeKind.Heading).Select(node => node.Title));
        Assert.Equal(changeDisplayPath ? "C > D" : "A > B", location.HeadingPath);
        Assert.Single(source.OcrCandidates);
        Assert.Empty(source.Blocks);
    }
}
