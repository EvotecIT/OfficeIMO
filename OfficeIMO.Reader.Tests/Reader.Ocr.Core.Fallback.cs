using OfficeIMO.Ocr;
using OfficeIMO.Reader;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderOcrCoreTests {
    [Theory]
    [InlineData("results")]
    [InlineData("async")]
    [InlineData("processor")]
    [InlineData("tree")]
    public async Task OcrPreservesChunkOnlyNativeContentThroughTransport(string route) {
        var source = CreateDocument(route == "tree" ? 0 : 1);
        var nativeLocation = new ReaderLocation { Path = "scan.pdf", Page = 1, BlockAnchor = "native", SourceBlockKind = "paragraph", HeadingPath = "A > B" };
        ReaderHeadingPath.SetHierarchyPath(nativeLocation, ReaderHeadingPath.Combine(new[] { "A > B" }));
        source.Blocks = new[] { new OfficeDocumentBlock { Id = "placeholder", Text = "", Location = new ReaderLocation { Path = "scan.pdf", Page = 1 } } };
        source.Chunks = new[] { new ReaderChunk { Id = "native", Text = "Native body", Markdown = "Native body", Location = nativeLocation,
            Warnings = new[] { "Original warning" }, Tables = new[] { new ReaderTable { Columns = new[] { "Code" }, Rows = new[] { (IReadOnlyList<string>)new[] { "A-100" } } } } } };
        source.Markdown = "Native body";
        if (route == "tree") {
            var child = CreateDocument(1);
            child.Chunks = new[] { new ReaderChunk { Id = "child-native", Text = "Child native", Location = new ReaderLocation { Path = "scan.pdf", Page = 1 } } };
            source.NestedDocuments = new[] { new OfficeDocumentNestedResult { Path = "child.pdf", Document = child } };
            source.Chunks = source.Chunks.Concat(new[] { new ReaderChunk { Id = "child-projection", Text = "Child native", Location = new ReaderLocation { Path = "scan.pdf!/child.pdf", Page = 1 } } }).ToArray();
        }
        source = OfficeDocumentReadResultJson.Deserialize(source.ToJson());
        var candidateOwner = route == "tree" ? source.NestedDocuments[0].Document : source;
        candidateOwner.Assets[0].PayloadBytes = new byte[] { 1 }; // Reader JSON intentionally omits binary assets.
        string before = source.ToJson();
        var engine = new DelegateOcrEngine("fallback", (_, _) => Task.FromResult(new OcrResult { Text = "Scanned amount 42" }));
        OfficeDocumentReadResult enriched;
        if (route == "results") enriched = source.ApplyOcrResults(new[] { new OfficeDocumentOcrTextResult { CandidateId = "ocr-1", Text = "Scanned amount 42" } }).Document;
        else if (route == "tree") enriched = (await source.ApplyOcrTreeAsync(engine)).Document;
        else if (route == "processor") {
            var reader = new OfficeDocumentReaderBuilder().AddProcessor(new OfficeDocumentOcrProcessor(engine)).Build();
            enriched = (await reader.ProcessDocumentAsync(source)).Document;
        } else enriched = (await source.ApplyOcrAsync(engine)).Document;
        enriched = OfficeDocumentReadResultJson.Deserialize(enriched.ToJson());

        var content = enriched.EnumerateContent().ToArray();
        var native = Assert.Single(content, item => item.Block?.Text == "Native body" || item.Chunk?.Text == "Native body");
        Assert.Single(content, item => item.Block?.Text == "Scanned amount 42" || item.Chunk?.Text == "Scanned amount 42");
        if (route == "tree") Assert.Single(content, item => item.Block?.Text == "Child native" || item.Chunk?.Text == "Child native");
        Assert.Equal("scan.pdf", native.Location!.Path);
        Assert.Equal(ReaderHeadingPath.Combine(new[] { "A > B" }), native.Location.HierarchyHeadingPath);
        Assert.Equal("A-100", Assert.Single(enriched.EnumerateTables()).Rows[0][0]);
        Assert.Equal("Original warning", Assert.Single(enriched.Chunks[0].Warnings!));
        Assert.Equal(before, source.ToJson());
    }
}
