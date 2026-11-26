using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderFb2Tests {
    [Fact]
    public void ImportedChaptersRetainSourceKindAndLogicalPageProvenance() {
        string path = Path.Combine(AppContext.BaseDirectory, "EbookFixtures", "harbour.fb2.zip");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();

        OfficeDocumentReadResult result = reader.ReadDocument(path);

        Assert.Equal(ReaderInputKind.Fb2, result.Kind);
        Assert.NotEmpty(result.Pages);
        Assert.Equal(OfficeDocumentPageProvenance.LogicalContainer, result.GetPageProvenance());
        Assert.NotEmpty(result.Chunks);
        Assert.All(result.Chunks, chunk => Assert.Equal(ReaderInputKind.Fb2, chunk.Kind));
    }
}
