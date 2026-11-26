using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderMobiTests {
    [Fact]
    public void ImportedChaptersRetainSourceKindAndLogicalPageProvenance() {
        string path = Path.Combine(AppContext.BaseDirectory, "EbookFixtures", "harbour.azw3");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(path);

        Assert.Equal(ReaderInputKind.Mobi, result.Kind);
        Assert.NotEmpty(result.Pages);
        Assert.Equal(OfficeDocumentPageProvenance.LogicalContainer, result.GetPageProvenance());
        Assert.NotEmpty(result.Chunks);
        Assert.All(result.Chunks, chunk => Assert.Equal(ReaderInputKind.Mobi, chunk.Kind));
    }
}
