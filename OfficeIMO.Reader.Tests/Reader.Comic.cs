using OfficeIMO.Archives;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderComicTests {
    [Theory]
    [InlineData("harbour.cbz", "officeimo.reader.cbz")]
    [InlineData("harbour.cb7", "officeimo.reader.cb7")]
    [InlineData("harbour.cbr", "officeimo.reader.cbr")]
    public void ImportedChaptersRetainSourceKindAndLogicalPageProvenance(string name, string handler) {
        string path = Path.Combine(AppContext.BaseDirectory, "EbookFixtures", name);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();

        using var source = File.OpenRead(path);
        OfficeDocumentReadResult result = reader.ReadDocument(source, Path.GetFileName(path), new ReaderOptions { ComputeHashes = true });

        Assert.Equal(ReaderInputKind.Comic, result.Kind);
        Assert.Contains(handler, result.CapabilitiesUsed);
        Assert.NotEmpty(result.Pages);
        Assert.Equal(OfficeDocumentPageProvenance.LogicalContainer, result.GetPageProvenance());
        Assert.NotEmpty(result.Chunks);
        Assert.False(string.IsNullOrWhiteSpace(result.Source.SourceHash));
        Assert.All(result.Chunks, chunk => {
            Assert.Equal(ReaderInputKind.Comic, chunk.Kind);
            Assert.Equal(result.Source.SourceHash, chunk.SourceHash);
            Assert.Equal(DocumentReaderEngine.ComputeChunkHash(chunk), chunk.ChunkHash);
        });
    }
    [Fact]
    public void EncryptedRarComicRetainsTheReaderEncryptionErrorCode() {
        string path = Path.Combine(AppContext.BaseDirectory, "EbookFixtures", "encrypted-headers.cbr");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
        var error = Assert.Throws<Rar5EncryptionNotSupportedException>(() => reader.ReadDocument(path));
        Assert.Equal(ReaderErrorCodes.UnsupportedEncryption, reader.ClassifyReadError(error, path));
    }

}
