using OfficeIMO.AsciiDoc;
using OfficeIMO.Reader;
using OfficeIMO.Reader.AsciiDoc;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderAsciiDocModularTests {
    [Fact]
    public void ReadAsciiDocDocument_EmitsTypedBlockChunksWithSourceLines() {
        const string source = "= Guide\n\n== Start\nParagraph\n\n* one\n** nested\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;

        ReaderChunk[] chunks = DocumentReaderAsciiDocExtensions.ReadAsciiDocDocument(document, "guide.adoc").ToArray();

        Assert.Equal(4, chunks.Length);
        Assert.All(chunks, chunk => Assert.Equal(ReaderInputKind.AsciiDoc, chunk.Kind));
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "heading" && chunk.Text == "Guide" && chunk.Location.StartLine == 1);
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "paragraph" && chunk.Text == "Paragraph");
        ReaderChunk list = Assert.Single(chunks, chunk => chunk.Location.SourceBlockKind == "unordered-list");
        Assert.Contains("nested", list.Markdown, StringComparison.Ordinal);
        Assert.Equal("Guide > Start", list.Location.HeadingPath);
    }

    [Fact]
    public void RegisteredHandler_DispatchesStreamAndReaderAddsHashes() {
        try {
            DocumentReaderAsciiDocRegistrationExtensions.RegisterAsciiDocHandler();
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes("= Registry\n\nContent\n"), writable: false);

            ReaderChunk[] chunks = DocumentReader.Read(stream, "registry.adoc").ToArray();

            Assert.NotEmpty(chunks);
            Assert.Equal(ReaderInputKind.AsciiDoc, DocumentReader.DetectKind("registry.adoc"));
            Assert.All(chunks, chunk => {
                Assert.Equal(ReaderInputKind.AsciiDoc, chunk.Kind);
                Assert.False(string.IsNullOrWhiteSpace(chunk.SourceId));
                Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash));
                Assert.Equal("asciidoc", chunk.Diagnostics?.SourceKind);
            });
        } finally {
            DocumentReaderAsciiDocRegistrationExtensions.UnregisterAsciiDocHandler();
        }
    }

    [Fact]
    public void ParserRecoveryDiagnostic_IsExposedAsReaderWarning() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("----\nunterminated"), writable: false);

        ReaderChunk chunk = Assert.Single(DocumentReaderAsciiDocExtensions.ReadAsciiDoc(stream, "broken.adoc"));

        Assert.NotNull(chunk.Warnings);
        Assert.Contains(chunk.Warnings!, warning => warning.StartsWith("ADOC001:", StringComparison.Ordinal));
    }

    [Fact]
    public void NonSeekableStream_EnforcesReaderInputLimit() {
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes("= Too much content for this limit\n"));

        IOException exception = Assert.Throws<IOException>(() => DocumentReaderAsciiDocExtensions.ReadAsciiDoc(
            stream,
            "limited.adoc",
            new ReaderOptions { MaxInputBytes = 8 }).ToArray());

        Assert.Contains("Input exceeds MaxInputBytes", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Phase1Blocks_ProduceSemanticChunksWithoutMetadataOrCompoundDuplicates() {
        const string source =
            "[.wide]\nTerm:: Definition\n\n" +
            "WARNING: Careful\n\n" +
            "* item\n+\nattached\n\n" +
            "[cols=2*]\n|===\n|A |B\n|===\n";

        ReaderChunk[] chunks = DocumentReaderAsciiDocExtensions.ReadAsciiDocDocument(
            AsciiDocDocument.Parse(source).Document,
            "phase1.adoc").ToArray();

        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "description-list" && chunk.Text.Contains("Term: Definition", StringComparison.Ordinal));
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "admonition" && chunk.Text == "WARNING: Careful");
        ReaderChunk list = Assert.Single(chunks, chunk => chunk.Location.SourceBlockKind == "unordered-list");
        Assert.Contains("attached", list.Markdown, StringComparison.Ordinal);
        Assert.Contains("attached", list.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(chunks, chunk => chunk.Location.SourceBlockKind == "raw");
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "table" && chunk.Text.Contains("A\tB", StringComparison.Ordinal));
    }

    [Fact]
    public void BlockChunks_ResolveDocumentAttributes() {
        const string source = ":product: OfficeIMO\n\nUse {product}.\n";

        ReaderChunk paragraph = Assert.Single(DocumentReaderAsciiDocExtensions.ReadAsciiDocDocument(
            AsciiDocDocument.Parse(source).Document,
            "attributes.adoc"));

        Assert.Contains("OfficeIMO", paragraph.Markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("{product}", paragraph.Markdown, StringComparison.Ordinal);
        Assert.DoesNotContain(paragraph.Warnings ?? Array.Empty<string>(), warning => warning.StartsWith("ADOCMD101:", StringComparison.Ordinal));
    }
}
