using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderMobiTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PalmDatabaseHandlerCanCoexistWithTheDefaultMobiPreset(bool registerFirst) {
        var custom = new ReaderHandlerRegistration {
            Id = "application.contacts",
            Kind = ReaderInputKind.Text,
            Extensions = new[] { ".pdb" },
            ReadStream = (_, _, _, _) => new[] { new ReaderChunk { Text = "Application contacts" } }
        };
        var builder = new OfficeDocumentReaderBuilder();
        if (registerFirst) builder.AddHandler(custom);
        builder.AddAllOfficeIMOHandlers();
        OfficeDocumentReader? fallbackReader = registerFirst ? null : builder.Build();
        if (!registerFirst) builder.AddHandler(custom);
        OfficeDocumentReader reader = builder.Build();

        Assert.Equal("Application contacts", Assert.Single(reader.ReadDocument(new byte[] { 65 }, "contacts.pdb").Chunks).Text);
        string azw3 = Path.Combine(AppContext.BaseDirectory, "EbookFixtures", "harbour.azw3");
        Assert.Equal(ReaderInputKind.Mobi, reader.ReadDocument(azw3).Kind);

        if (fallbackReader != null) {
            string pdb = Path.Combine(AppContext.BaseDirectory, "EbookFixtures", "harbour.pdb");
            OfficeDocumentReadResult original = fallbackReader.ReadDocument(pdb);
            Assert.Equal(ReaderInputKind.Mobi, original.Kind);
            Assert.Contains("officeimo.reader.palmdoc", original.CapabilitiesUsed);
            Assert.Contains(original.Chunks, chunk => chunk.Text.IndexOf("Signal received.", StringComparison.Ordinal) >= 0);
        }
    }

    [Fact]
    public void ImportedChaptersRetainSourceKindAndLogicalPageProvenance() {
        string path = Path.Combine(AppContext.BaseDirectory, "EbookFixtures", "harbour.azw3");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();

        using var source = File.OpenRead(path);
        OfficeDocumentReadResult result = reader.ReadDocument(source, Path.GetFileName(path), new ReaderOptions { ComputeHashes = true });

        Assert.Equal(ReaderInputKind.Mobi, result.Kind);
        Assert.NotEmpty(result.Pages);
        Assert.Equal(OfficeDocumentPageProvenance.LogicalContainer, result.GetPageProvenance());
        Assert.NotEmpty(result.Chunks);
        Assert.False(string.IsNullOrWhiteSpace(result.Source.SourceHash));
        Assert.All(result.Chunks, chunk => {
            Assert.Equal(ReaderInputKind.Mobi, chunk.Kind);
            Assert.Equal(result.Source.SourceHash, chunk.SourceHash);
            Assert.Equal(DocumentReaderEngine.ComputeChunkHash(chunk), chunk.ChunkHash);
        });
    }
}
