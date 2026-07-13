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
        ReaderChunk[] chunks = LatexReaderAdapter.Read(
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
    public void BuilderHandler_DispatchesTexStreamAndAddsHashes() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddLatexHandler().Build();
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(Source), writable: false);

        ReaderChunk[] chunks = reader.Read(stream, "guide.tex").ToArray();

        Assert.NotEmpty(chunks);
        Assert.Equal(ReaderInputKind.Latex, reader.DetectKind("guide.tex"));
        Assert.All(chunks, chunk => {
            Assert.Equal(ReaderInputKind.Latex, chunk.Kind);
            Assert.False(string.IsNullOrWhiteSpace(chunk.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash));
        });
    }

    [Fact]
    public void UnrecognizedPlainTexProfile_IsPreservedAndWarned() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Plain \\hbox{TeX}"), writable: false);

        ReaderChunk[] chunks = LatexReaderAdapter.Read(stream, "plain.tex").ToArray();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()), warning => warning.StartsWith("LATEXR001:", StringComparison.Ordinal));
    }

    [Fact]
    public void WholeDocumentPlainTex_EmitsFallbackChunk() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Plain \\hbox{TeX}"), writable: false);

        ReaderChunk chunk = Assert.Single(LatexReaderAdapter.Read(
            stream,
            "plain.tex",
            latexOptions: new ReaderLatexOptions { ChunkByBlock = false }));

        Assert.Contains("Plain \\hbox{TeX}", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("```latex", chunk.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void NonSeekableStream_EnforcesInputLimit() {
        using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(Source));

        IOException exception = Assert.Throws<IOException>(() => LatexReaderAdapter.Read(
            stream,
            "limited.tex",
            new ReaderOptions { MaxInputBytes = 8 }).ToArray());

        Assert.Contains("Input exceeds MaxInputBytes", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void WholeDocumentChunk_TextIncludesAllProjectedSemanticBlocks() {
        ReaderChunk chunk = Assert.Single(LatexReaderAdapter.Read(
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

        ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.Parse(source).Document).ToArray();

        ReaderChunk definitions = Assert.Single(chunks, static chunk => chunk.Location.SourceBlockKind == "list-description");
        Assert.Contains("Term: Definition", definitions.Text, StringComparison.Ordinal);
        Assert.Contains("Term", definitions.Markdown, StringComparison.Ordinal);
        ReaderChunk figure = Assert.Single(chunks, static chunk => chunk.Location.SourceBlockKind == "figure");
        Assert.Contains("Plot caption", figure.Text, StringComparison.Ordinal);
        Assert.Contains("Plot caption", figure.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void FloatingTableChunk_IncludesWrapperCaptionAndLabelMetadata() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\begin{table}\\caption{Important values}\\label{tab:values}" +
            "\\begin{tabular}{ll}A & B\\\\\\end{tabular}\\end{table}\n" +
            "\\end{document}\n";

        ReaderChunk table = Assert.Single(LatexReaderAdapter.Read(
            LatexDocument.Parse(source).Document,
            "table.tex"), static chunk => chunk.Location.SourceBlockKind == "table");

        Assert.Contains("Important values", table.Text, StringComparison.Ordinal);
        Assert.Contains("caption=\"Important values\"", table.Markdown, StringComparison.Ordinal);
        Assert.Contains("#tab:values", table.Markdown, StringComparison.Ordinal);
        Assert.Equal(3, table.Location.StartLine);
    }
}
