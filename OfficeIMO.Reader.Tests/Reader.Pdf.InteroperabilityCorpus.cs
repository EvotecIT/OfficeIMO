using System.Text.Json;
using System.Threading;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderPdfInteroperabilityCorpusTests {
    [Fact]
    public void ReaderPdf_ProjectsPinnedProducerCorpusWithDeterministicProvenanceAndHierarchy() {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(InteroperabilityRoot, "corpus-manifest.json")));
        foreach (JsonElement item in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            string id = RequireString(item, "id");
            string fileName = RequireString(item, "file");
            byte[] bytes = File.ReadAllBytes(Path.Combine(InteroperabilityRoot, fileName));
            var readerOptions = new ReaderOptions {
                ComputeHashes = true,
                MaxChars = 4_096,
                MaxInputBytes = bytes.LongLength
            };

            OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(bytes, fileName, readerOptions);
            OfficeDocumentReadResult repeated = PdfReaderAdapter.ReadDocument(bytes, fileName, readerOptions);

            Assert.Equal(ReaderInputKind.Pdf, result.Kind);
            Assert.Equal(item.GetProperty("pageCount").GetInt32(), result.Pages.Count);
            Assert.Equal(RequireString(item, "sha256"), result.Source.SourceHash);
            Assert.Equal(bytes.LongLength, result.Source.LengthBytes);
            Assert.Equal(result.Chunks.Select(chunk => chunk.Id), repeated.Chunks.Select(chunk => chunk.Id));
            Assert.Equal(result.Chunks.Select(chunk => chunk.ChunkHash), repeated.Chunks.Select(chunk => chunk.ChunkHash));
            Assert.NotEmpty(result.Chunks);
            Assert.All(result.Pages, page => {
                Assert.InRange(page.Number ?? 0, 1, result.Pages.Count);
                Assert.Equal(fileName, page.Location.Path);
            });
            Assert.All(result.Chunks, chunk => {
                Assert.Equal(ReaderInputKind.Pdf, chunk.Kind);
                Assert.Equal(fileName, chunk.Location.Path);
                Assert.InRange(chunk.Location.Page ?? 0, 1, result.Pages.Count);
                Assert.Equal(result.Source.SourceHash, chunk.SourceHash);
                Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash));
                Assert.NotNull(chunk.Diagnostics);
                Assert.Equal("pdf", chunk.Diagnostics!.SourceKind);
                Assert.Equal(result.Pages.Count, chunk.Diagnostics.PageCount);
            });
            Assert.True(result.Chunks.Select(chunk => chunk.Location.Page).SequenceEqual(
                result.Chunks.Select(chunk => chunk.Location.Page).OrderBy(page => page)),
                id + " emitted chunks outside source page order.");
            Assert.All(result.Blocks, block => Assert.InRange(block.Location.Page ?? 0, 1, result.Pages.Count));
            Assert.All(result.Assets, asset => Assert.Equal(fileName, asset.Location.Path));
            Assert.Equal(result.Assets.Count, result.Pages.Sum(page => page.Assets.Count));

            var hierarchyOptions = new ReaderHierarchicalChunkingOptions {
                MaxTokens = 256,
                OverlapTokens = 24,
                MaxInputChunks = 128,
                MaxOutputChunks = 512,
                MaxHierarchyDepth = 8
            };
            ReaderChunkHierarchyResult hierarchy = ReaderHierarchicalChunker.Chunk(result, hierarchyOptions);
            ReaderChunkHierarchyResult repeatedHierarchy = ReaderHierarchicalChunker.Chunk(repeated, hierarchyOptions);
            Assert.Equal(ReaderChunkHierarchySchema.Id, hierarchy.SchemaId);
            Assert.Equal(hierarchy.ToJson(), repeatedHierarchy.ToJson());
            Assert.Equal(hierarchy.Chunks.Count, hierarchy.Segments.Count);
            Assert.Contains(hierarchy.Nodes, node => node.Kind == ReaderChunkHierarchyNodeKind.Document && node.Id == hierarchy.RootNodeId);
            Assert.Equal(result.Pages.Count, hierarchy.Nodes.Count(node => node.Kind == ReaderChunkHierarchyNodeKind.Container));
            Assert.All(hierarchy.Segments, segment => {
                Assert.True(segment.EndCharacter > segment.StartCharacter);
                Assert.InRange(segment.TokenCount, 1, hierarchyOptions.MaxTokens);
            });
        }
    }

    [Theory]
    [InlineData("microsoft-excel-16.109-native-excel-daily-workbook.pdf", 2, 1, 1)]
    [InlineData("microsoft-powerpoint-16.109-native-powerpoint-dense-layout.pdf", 1, 0, 1)]
    [InlineData("microsoft-word-16.109-native-word-report.pdf", 1, 0, 1)]
    [InlineData("microsoft-word-windows-word-business-delivery-summary.pdf", 9, 14, 0)]
    public void ReaderPdf_ReconstructsNativeProducerTablesAndVisualsWithBoundedConfidence(
        string fileName,
        int expectedPages,
        int minimumTables,
        int minimumAssets) {
        string path = Path.Combine(ReferenceBaselineRoot, fileName);
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            path,
            new ReaderOptions { ComputeHashes = true, MaxChars = 8_000, MaxInputBytes = new FileInfo(path).Length });

        Assert.Equal(expectedPages, result.Pages.Count);
        Assert.True(
            result.Tables.Count >= minimumTables,
            $"{fileName} reconstructed {result.Tables.Count} tables; expected at least {minimumTables}.");
        if (fileName == "microsoft-word-windows-word-business-delivery-summary.pdf") {
            Assert.Equal(11, result.Tables.Count(table =>
                table.Columns.Count >= 2 &&
                NormalizeTableLabel(table.Columns[0]) == "deliveryworksheet" &&
                NormalizeTableLabel(table.Columns[1]) == "response"));
            Assert.Contains(result.Tables, table =>
                table.Columns.Select(NormalizeTableLabel).SequenceEqual(new[] { "role", "name", "decision", "date", "notes" }));
        }
        Assert.True(result.Assets.Count >= minimumAssets);
        Assert.Equal(result.Assets.Count, result.Visuals.Count);
        Assert.All(result.Tables, table => {
            Assert.NotNull(table.Diagnostics);
            Assert.InRange(table.Diagnostics!.Confidence, 0D, 1D);
            Assert.InRange(table.Diagnostics.SchemaConfidence, 0D, 1D);
            Assert.InRange(table.Diagnostics.ColumnGeometryConfidence, 0D, 1D);
            Assert.NotNull(table.Location);
            Assert.InRange(table.Location!.Page ?? 0, 1, expectedPages);
        });
        Assert.All(result.Assets, asset => {
            Assert.Equal(result.Source.Path, asset.Location.Path);
            Assert.EndsWith(fileName, asset.Location.Path, StringComparison.OrdinalIgnoreCase);
            Assert.InRange(asset.Location.Page ?? 0, 1, expectedPages);
        });
    }

    private static string NormalizeTableLabel(string value) =>
        new string(value.Where(static character => !char.IsWhiteSpace(character)).ToArray()).ToLowerInvariant();

    [Fact]
    public void ReaderPdf_RejectsCorpusInputBeyondBudgetAndHonorsPreCancellation() {
        string path = Path.Combine(InteroperabilityRoot, "verapdf-devicen-content.pdf");
        byte[] bytes = File.ReadAllBytes(path);
        Assert.Throws<IOException>(() => PdfReaderAdapter.ReadDocument(
            bytes,
            Path.GetFileName(path),
            new ReaderOptions { MaxInputBytes = bytes.LongLength - 1 }));

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => {
            _ = PdfReaderAdapter.ReadDocument(
                bytes,
                Path.GetFileName(path),
                cancellationToken: cancellation.Token);
        });
    }

    private static string RequireString(JsonElement element, string propertyName) =>
        element.GetProperty(propertyName).GetString() ?? throw new InvalidDataException("Missing " + propertyName + ".");

    private static string InteroperabilityRoot => Path.Combine(AppContext.BaseDirectory, "PdfInteroperability");
    private static string ReferenceBaselineRoot => Path.Combine(AppContext.BaseDirectory, "PdfReferenceBaselines");
}
