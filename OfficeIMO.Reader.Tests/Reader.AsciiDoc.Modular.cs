using OfficeIMO.AsciiDoc;
using OfficeIMO.Reader;
using OfficeIMO.Reader.AsciiDoc;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderAsciiDocModularTests {
    [Fact]
    public void BlockChunksResolveExternalFootnoteDefinitionsAndHeadingReferences() {
        var document = AsciiDocDocument.Parse("[[section]]\n== A section\n\nFirst footnote:note[Definition].\n\nSecond footnote:note[] xref:section[].\n");
        var options = new ReaderAsciiDocOptions();
        ReaderChunk[] chunks = AsciiDocReaderAdapter.Read(document, "references.adoc", asciiDocOptions: options).ToArray();
        ReaderChunk second = Assert.Single(chunks, chunk => chunk.Text.StartsWith("Second", StringComparison.Ordinal));
        Assert.Contains("Definition", second.Markdown);
        Assert.Contains("[A section](#section)", second.Markdown);
        Assert.DoesNotContain(second.Warnings ?? Array.Empty<string>(), warning => warning.StartsWith("ADOCREF004:", StringComparison.Ordinal) || warning.StartsWith("ADOCMD104:", StringComparison.Ordinal));
        Assert.Null(options.MarkdownOptions.References);
    }
    [Fact]
    public void ReadAsciiDocDocument_EmitsTypedBlockChunksWithSourceLines() {
        const string source = "= Guide\n\n== Start\nParagraph\n\n* one\n** nested\n";
        AsciiDocDocument document = AsciiDocDocument.ParseResult(source).Document;

        ReaderChunk[] chunks = AsciiDocReaderAdapter.Read(document, "guide.adoc").ToArray();

        Assert.Equal(4, chunks.Length);
        Assert.All(chunks, chunk => Assert.Equal(ReaderInputKind.AsciiDoc, chunk.Kind));
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "heading" && chunk.Text == "Guide" && chunk.Location.StartLine == 1);
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "paragraph" && chunk.Text == "Paragraph");
        ReaderChunk list = Assert.Single(chunks, chunk => chunk.Location.SourceBlockKind == "unordered-list");
        Assert.Contains("nested", list.Markdown, StringComparison.Ordinal);
        Assert.Equal("Guide > Start", list.Location.HeadingPath);
    }

    [Fact]
    public void BuilderHandler_DispatchesStreamAndReaderAddsHashes() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAsciiDocHandler().Build();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("= Registry\n\nContent\n"), writable: false);

        ReaderChunk[] chunks = reader.Read(stream, "registry.adoc").ToArray();

        Assert.NotEmpty(chunks);
        Assert.Equal(ReaderInputKind.AsciiDoc, reader.DetectKind("registry.adoc"));
        Assert.All(chunks, chunk => {
            Assert.Equal(ReaderInputKind.AsciiDoc, chunk.Kind);
            Assert.False(string.IsNullOrWhiteSpace(chunk.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash));
            Assert.Equal("asciidoc", chunk.Diagnostics?.SourceKind);
        });
    }

    [Fact]
    public void ParserRecoveryDiagnostic_IsExposedAsReaderWarning() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("----\nunterminated"), writable: false);

        ReaderChunk chunk = Assert.Single(AsciiDocReaderAdapter.Read(stream, "broken.adoc"));

        Assert.NotNull(chunk.Warnings);
        Assert.Contains(chunk.Warnings!, warning => warning.StartsWith("ADOC001:", StringComparison.Ordinal));
    }

    [Fact]
    public void NonSeekableStream_EnforcesReaderInputLimit() {
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes("= Too much content for this limit\n"));

        IOException exception = Assert.Throws<IOException>(() => AsciiDocReaderAdapter.Read(
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

        ReaderChunk[] chunks = AsciiDocReaderAdapter.Read(
            AsciiDocDocument.ParseResult(source).Document,
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

        ReaderChunk paragraph = Assert.Single(AsciiDocReaderAdapter.Read(
            AsciiDocDocument.ParseResult(source).Document,
            "attributes.adoc"));

        Assert.Contains("OfficeIMO", paragraph.Markdown, StringComparison.Ordinal);
        Assert.Equal("Use OfficeIMO.", paragraph.Text);
        Assert.DoesNotContain("{product}", paragraph.Markdown, StringComparison.Ordinal);
        Assert.DoesNotContain(paragraph.Warnings ?? Array.Empty<string>(), warning => warning.StartsWith("ADOCMD101:", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SourceOrderedAttributeChangesResolveInTextAndMarkdown(bool chunkByBlock) {
        const string source = ":color: blue\n\nFirst {color}.\n\n:color: red\n\nSecond {color}.\n";
        ReaderChunk[] chunks = AsciiDocReaderAdapter.Read(AsciiDocDocument.Parse(source), "colors.adoc",
            asciiDocOptions: new ReaderAsciiDocOptions { ChunkByBlock = chunkByBlock }).ToArray();
        string text = string.Join("\n", chunks.Select(chunk => chunk.Text));
        string markdown = string.Join("\n", chunks.Select(chunk => chunk.Markdown));
        Assert.Contains("First blue.", text);
        Assert.Contains("Second red.", text);
        Assert.Contains("First blue.", markdown);
        Assert.Contains("Second red.", markdown);
    }

    [Fact]
    public void ReaderForwardsTheConfiguredCompoundDepthLimit() {
        AsciiDocDocument document = AsciiDocDocument.Parse("====\nParagraph\n====\n");
        var options = new ReaderAsciiDocOptions(); options.MarkdownOptions.MaximumBlockNestingDepth = 1;
        Assert.Throws<InvalidDataException>(() => AsciiDocReaderAdapter.Read(document, "depth.adoc", asciiDocOptions: options).ToArray());
    }

    [Fact]
    public void StreamReaderHonorsTheNativeCharacterLimitAndRestoresCallerPosition() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Text longer than eight characters."), writable: false);
        var options = new ReaderAsciiDocOptions(); options.ParseOptions.MaximumInputLength = 8;
        Assert.Throws<InvalidDataException>(() => AsciiDocReaderAdapter.Read(stream, "limit.adoc", asciiDocOptions: options).ToArray());
        Assert.True(stream.CanRead);
        Assert.Equal(0, stream.Position);
    }
}
