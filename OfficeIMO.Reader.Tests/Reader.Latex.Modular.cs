using System.Threading;
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
            LatexDocument.ParseResult(Source).Document,
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
            LatexDocument.ParseResult(Source).Document,
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

        ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.ParseResult(source).Document).ToArray();

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
            LatexDocument.ParseResult(source).Document,
            "table.tex"), static chunk => chunk.Location.SourceBlockKind == "table");

        Assert.Contains("Important values", table.Text, StringComparison.Ordinal);
        Assert.Contains("caption=\"Important values\"", table.Markdown, StringComparison.Ordinal);
        Assert.Contains("#tab:values", table.Markdown, StringComparison.Ordinal);
        Assert.Equal(3, table.Location.StartLine);
    }
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MixedProjectionRetainsFallbacksCodeQuotesAndFormattedTableText(bool blocks) {
        const string source = "\\documentclass{article}\\begin{document}Before.\n" +
            "\\begin{custom}IMPORTANT UNKNOWN\\end{custom}\n" +
            "\\begin{verbatim}IMPORTANT CODE\\end{verbatim}\n" +
            "\\begin{quote}IMPORTANT \\textbf{QUOTE}\\begin{comment}PRIVATE\\end{comment}\\end{quote}\n" +
            "\\begin{tabular}{ll}\\textbf{Bold}&B\\\\\\end{tabular}\nAfter.\\end{document}";
        ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.Parse(source), latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks }).ToArray();
        string text = string.Join("\n", chunks.Select(static chunk => chunk.Text));
        string markdown = string.Join("\n", chunks.Select(static chunk => chunk.Markdown));
        foreach (string value in new[] { "Before", "IMPORTANT UNKNOWN", "IMPORTANT CODE", "IMPORTANT QUOTE", "Bold\tB", "After" }) Assert.Contains(value, text, StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE", text, StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE", markdown, StringComparison.Ordinal);
        Assert.Contains(chunks.SelectMany(static chunk => chunk.Warnings ?? Array.Empty<string>()), static warning => warning.StartsWith("LATEXMD299:", StringComparison.Ordinal));
    }

    [Fact]
    public void ParsedDocumentEditsFlowIntoReaderTextAndMarkdown() {
        LatexDocument document = LatexDocument.Parse(Source);
        document.Headings[0].Title = "Updated title";
        document.Paragraphs[0].Content = "Updated paragraph";
        ReaderChunk[] chunks = LatexReaderAdapter.Read(document).ToArray();
        Assert.Contains(chunks, static chunk => chunk.Text == "Updated title");
        Assert.Contains(chunks, static chunk => chunk.Text == "Updated paragraph" && chunk.Markdown?.Contains("Updated paragraph", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void RegistrationSnapshotsEveryNativeLimitAndCustomOpaqueEnvironment() {
        var parse = new LatexParseOptions {
            MaximumInputBytes = 8, MaximumExpansionInputLength = 123, MaximumExpansionTokenCount = 7
        };
        parse.VerbatimEnvironmentNames.Add("opaque");
        ReaderLatexOptions clone = ReaderLatexOptionsCloner.Clone(new ReaderLatexOptions { ParseOptions = parse });
        Assert.Equal(8, clone.ParseOptions.MaximumInputBytes);
        Assert.Equal(123, clone.ParseOptions.MaximumExpansionInputLength);
        Assert.Equal(7, clone.ParseOptions.MaximumExpansionTokenCount);
        parse.VerbatimEnvironmentNames.Clear();
        Assert.Contains("opaque", clone.ParseOptions.VerbatimEnvironmentNames);
        using var limited = new MemoryStream(Encoding.UTF8.GetBytes(Source));
        Assert.ThrowsAny<IOException>(() => LatexReaderAdapter.Read(limited, latexOptions: clone).ToArray());
        clone.ParseOptions.MaximumInputBytes = null;
        const string opaque = "\\documentclass{article}\\begin{document}\\begin{opaque}\\section{FAKE}\\end{opaque}\\end{document}";
        using var input = new MemoryStream(Encoding.UTF8.GetBytes(opaque));
        ReaderChunk[] chunks = LatexReaderAdapter.Read(input, latexOptions: clone).ToArray();
        Assert.DoesNotContain(chunks, static chunk => chunk.Location.SourceBlockKind == "heading");
        Assert.Contains(chunks, static chunk => chunk.Location.SourceBlockKind == "verbatim" && chunk.Text.Contains("FAKE", StringComparison.Ordinal));
    }

    [Fact]
    public void ReaderLoadingHonorsCancellationAndPreservesCallerStreamState() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(Source));
        stream.Position = 5;
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => LatexReaderAdapter.Read(stream, cancellationToken: cancellation.Token).ToArray());
        Assert.Equal(5, stream.Position);
        Assert.True(stream.CanRead);
        Assert.NotEmpty(LatexReaderAdapter.Read(stream));
        Assert.Equal(5, stream.Position);
    }
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OmittedOnlyContentStillProducesFidelityWarnings(bool blocks) {
        LatexDocument document = LatexDocument.Parse("\\begin{document}\\begin{custom}Important\\end{custom}\\end{document}");
        ReaderChunk[] chunks = LatexReaderAdapter.Read(document, latexOptions: new ReaderLatexOptions {
            ChunkByBlock = blocks,
            MarkdownOptions = new OfficeIMO.Latex.Markdown.LatexToMarkdownOptions { PreserveUnsupportedAsSource = false }
        }).ToArray();
        Assert.Contains(chunks.SelectMany(static chunk => chunk.Warnings ?? Array.Empty<string>()), static warning => warning.StartsWith("LATEXMD299:", StringComparison.Ordinal));
        Assert.All(chunks, static chunk => Assert.Equal(string.Empty, chunk.Text));
    }

}
