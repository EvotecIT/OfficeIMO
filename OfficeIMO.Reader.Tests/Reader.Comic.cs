using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderComicTests {
    [Fact]
    public void ImportedChaptersRetainSourceKindAndLogicalPageProvenance() {
        string path = Path.Combine(AppContext.BaseDirectory, "EbookFixtures", "harbour.cbz");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(path);

        Assert.Equal(ReaderInputKind.Comic, result.Kind);
        Assert.NotEmpty(result.Pages);
        Assert.Equal(OfficeDocumentPageProvenance.LogicalContainer, result.GetPageProvenance());
        Assert.NotEmpty(result.Chunks);
        Assert.All(result.Chunks, chunk => Assert.Equal(ReaderInputKind.Comic, chunk.Kind));
    }
}
