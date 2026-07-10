using OfficeIMO.Latex;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Latex;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderLatexModularTests {
    private const string Source =
        "\\documentclass{article}\n\\title{Guide}\n\\begin{document}\n\\maketitle\n" +
        "\\section{Start}\nParagraph with \\textbf{bold} and $x^2$.\n\n" +
        "\\begin{itemize}\n\\item One\n\\item Two\n\\end{itemize}\n" +
        "\\begin{tabular}{ll}\nA & B\\\\\nC & D\\\\\n\\end{tabular}\n" +
        "\\end{document}\n";

    [Fact]
    public void ParsedDocument_EmitsSemanticChunksWithHierarchyAndMathDiagnostics() {
        ReaderChunk[] chunks = DocumentReaderLatexExtensions.ReadLatexDocument(
            LatexDocument.Parse(Source).Document,
            "guide.tex").ToArray();

        Assert.All(chunks, chunk => Assert.Equal(ReaderInputKind.Latex, chunk.Kind));
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "title" && chunk.Text == "Guide");
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "heading" && chunk.Text == "Start");
        ReaderChunk paragraph = Assert.Single(chunks, chunk => chunk.Location.SourceBlockKind == "paragraph");
        Assert.Contains("bold", paragraph.Text, StringComparison.Ordinal);
        Assert.Contains(paragraph.Warnings ?? Array.Empty<string>(), warning => warning.StartsWith("LATEXMD101:", StringComparison.Ordinal));
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "list-unordered" && chunk.Text.Contains("One", StringComparison.Ordinal));
        Assert.Contains(chunks, chunk => chunk.Location.SourceBlockKind == "table" && chunk.Text.Contains("A\tB", StringComparison.Ordinal));
        Assert.Equal("Start", paragraph.Location.HeadingPath);
    }

    [Fact]
    public void RegisteredHandler_DispatchesTexStreamAndAddsHashes() {
        try {
            DocumentReaderLatexRegistrationExtensions.RegisterLatexHandler();
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(Source), writable: false);

            ReaderChunk[] chunks = DocumentReader.Read(stream, "guide.tex").ToArray();

            Assert.NotEmpty(chunks);
            Assert.Equal(ReaderInputKind.Latex, DocumentReader.DetectKind("guide.tex"));
            Assert.All(chunks, chunk => {
                Assert.Equal(ReaderInputKind.Latex, chunk.Kind);
                Assert.False(string.IsNullOrWhiteSpace(chunk.SourceId));
                Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash));
            });
        } finally {
            DocumentReaderLatexRegistrationExtensions.UnregisterLatexHandler();
        }
    }

    [Fact]
    public void UnrecognizedPlainTexProfile_IsPreservedAndWarned() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Plain \\hbox{TeX}"), writable: false);

        ReaderChunk[] chunks = DocumentReaderLatexExtensions.ReadLatex(stream, "plain.tex").ToArray();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()), warning => warning.StartsWith("LATEXR001:", StringComparison.Ordinal));
    }

    [Fact]
    public void NonSeekableStream_EnforcesInputLimit() {
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(Source));

        IOException exception = Assert.Throws<IOException>(() => DocumentReaderLatexExtensions.ReadLatex(
            stream,
            "limited.tex",
            new ReaderOptions { MaxInputBytes = 8 }).ToArray());

        Assert.Contains("Input exceeds MaxInputBytes", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void WholeDocumentChunk_TextIncludesAllProjectedSemanticBlocks() {
        ReaderChunk chunk = Assert.Single(DocumentReaderLatexExtensions.ReadLatexDocument(
            LatexDocument.Parse(Source).Document,
            "guide.tex",
            latexOptions: new ReaderLatexOptions { ChunkByBlock = false }));

        Assert.Contains("Guide", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("Start", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("Paragraph with bold", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("One", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("A\tB", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("# Guide", chunk.Markdown, StringComparison.Ordinal);
        Assert.Contains("- One", chunk.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void DescriptionListsAndFigureCaptions_ArePresentInBlockTextAndMarkdown() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\begin{description}\\item[Term] Definition\\end{description}\n" +
            "\\begin{figure}\\includegraphics{plot.png}\\caption{Plot caption}\\end{figure}\n" +
            "\\end{document}\n";

        ReaderChunk[] chunks = DocumentReaderLatexExtensions.ReadLatexDocument(LatexDocument.Parse(source).Document).ToArray();

        ReaderChunk definitions = Assert.Single(chunks, static chunk => chunk.Location.SourceBlockKind == "list-description");
        Assert.Contains("Term: Definition", definitions.Text, StringComparison.Ordinal);
        Assert.Contains("Term", definitions.Markdown, StringComparison.Ordinal);
        ReaderChunk figure = Assert.Single(chunks, static chunk => chunk.Location.SourceBlockKind == "figure");
        Assert.Contains("Plot caption", figure.Text, StringComparison.Ordinal);
        Assert.Contains("Plot caption", figure.Markdown, StringComparison.Ordinal);
    }
}
