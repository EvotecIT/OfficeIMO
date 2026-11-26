using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderFb2Tests {
    [Fact]
    public void ImportedChaptersRetainSourceKindAndLogicalPageProvenance() {
        string path = Path.Combine(AppContext.BaseDirectory, "EbookFixtures", "harbour.fb2.zip");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();

        using var source = File.OpenRead(path);
        OfficeDocumentReadResult result = reader.ReadDocument(source, Path.GetFileName(path), new ReaderOptions { ComputeHashes = true });

        Assert.Equal(ReaderInputKind.Fb2, result.Kind);
        Assert.NotEmpty(result.Pages);
        Assert.Equal(OfficeDocumentPageProvenance.LogicalContainer, result.GetPageProvenance());
        Assert.NotEmpty(result.Chunks);
        Assert.False(string.IsNullOrWhiteSpace(result.Source.SourceHash));
        Assert.All(result.Chunks, chunk => {
            Assert.Equal(ReaderInputKind.Fb2, chunk.Kind);
            Assert.Equal(result.Source.SourceHash, chunk.SourceHash);
            Assert.Equal(DocumentReaderEngine.ComputeChunkHash(chunk), chunk.ChunkHash);
        });
    }
}
