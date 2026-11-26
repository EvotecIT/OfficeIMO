using System.Text;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderMobiTests {
    [Theory]
    [InlineData(".pdb", true)]
    [InlineData(".pdb", false)]
    [InlineData(".prc", true)]
    [InlineData(".prc", false)]
    public void PalmDatabaseHandlerCanCoexistWithTheDefaultMobiPreset(string extension, bool registerFirst) {
        var custom = new ReaderHandlerRegistration {
            Id = "application.contacts",
            Kind = ReaderInputKind.Text,
            Extensions = new[] { extension },
            ReadStream = (_, _, _, _) => new[] { new ReaderChunk { Text = "Application contacts" } }
        };
        var builder = new OfficeDocumentReaderBuilder();
        if (registerFirst) builder.AddHandler(custom);
        builder.AddAllOfficeIMOHandlers();
        OfficeDocumentReader? fallbackReader = registerFirst ? null : builder.Build();
        if (!registerFirst) builder.AddHandler(custom);
        OfficeDocumentReader reader = builder.Build();

        Assert.Equal("Application contacts", Assert.Single(reader.ReadDocument(new byte[] { 65 }, "contacts" + extension).Chunks).Text);
        string fixtures = Path.Combine(AppContext.BaseDirectory, "EbookFixtures");
        string azw3 = Path.Combine(fixtures, "harbour.azw3");
        Assert.Equal(ReaderInputKind.Mobi, reader.ReadDocument(azw3).Kind);
        Assert.Equal(ReaderInputKind.Mobi, ReadPalmEbook(reader, extension == ".pdb" ? ".prc" : ".pdb").Kind);

        if (fallbackReader != null) {
            OfficeDocumentReadResult original = ReadPalmEbook(fallbackReader, extension);
            Assert.Equal(ReaderInputKind.Mobi, original.Kind);
            Assert.Contains(extension == ".pdb" ? "officeimo.reader.palmdoc" : "officeimo.reader.mobi-prc", original.CapabilitiesUsed);
            Assert.Contains(original.Chunks, chunk => chunk.Text.IndexOf("Signal received.", StringComparison.Ordinal) >= 0);
        }

        OfficeDocumentReadResult ReadPalmEbook(OfficeDocumentReader preset, string suffix) {
            string file = suffix == ".pdb" ? "harbour.pdb" : "harbour.mobi";
            return preset.ReadDocument(File.ReadAllBytes(Path.Combine(fixtures, file)), "harbour" + suffix);
        }
    }

    [Theory]
    [InlineData("Notes on %PDF- headers")]
    [InlineData("%PDF- headers")]
    [InlineData("ITSF notes")]
    [InlineData("MSCF notes")]
    public async Task DetectKeepsPalmDocIdentityWhenTheTitleResemblesAnotherFormat(string title) {
        string path = Path.Combine(AppContext.BaseDirectory, "EbookFixtures", "harbour.pdb");
        byte[] bytes = File.ReadAllBytes(path);
        Array.Clear(bytes, 0, 32);
        Encoding.ASCII.GetBytes(title).CopyTo(bytes, 0);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers().Build();
        using var input = new MemoryStream(bytes, writable: false);

        var options = new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent, InspectContainers = true };
        ReaderDetectionResult detected = reader.Detect(input, "renamed.pdf", options);
        ReaderDetectionResult detectedAsync = await reader.DetectAsync(input, "renamed.pdf", options);

        Assert.All(new[] { detected, detectedAsync }, result => {
            Assert.Equal(ReaderInputKind.Pdf, result.ExtensionKind);
            Assert.Equal(ReaderInputKind.Mobi, result.ContentKind);
            Assert.Equal(ReaderInputKind.Mobi, result.Kind);
            Assert.Contains("signature:palm-ebook", result.Evidence);
        });
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
