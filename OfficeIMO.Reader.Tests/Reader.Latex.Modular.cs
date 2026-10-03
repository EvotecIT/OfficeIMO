using System.Threading;
using OfficeIMO.Latex;
using OfficeIMO.Latex.Markdown;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Latex;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderLatexModularTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SourceAfterDocumentEndDoesNotLeakIntoReaderChunks(bool blocks) {
        const string source = @"\documentclass{article}\begin{document}Public\end{document}TRAILING\section{Inactive}";
        var chunks = LatexReaderAdapter.Read(LatexDocument.Parse(source), readerOptions: new ReaderOptions { MaxChars = 4 },
            latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks, IncludeDiagnostics = false }).ToArray();
        Assert.Equal("Public", string.Concat(chunks.Select(chunk => chunk.Text)));
        Assert.DoesNotContain(chunks, chunk => chunk.Markdown!.Contains("TRAILING"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SplitOpaqueContentRetainsItsCodeCarrierAndExactPayload(bool blocks) {
        const string payload = "  # Heading\n*value* [link](url) <a id=\"fake\"></a>\n```\n";
        string content = string.Concat(Enumerable.Repeat(payload, 6));
        LatexDocument document = LatexDocument.Parse("\\documentclass{article}\\begin{document}\\begin{verbatim}\n" + content + "\\end{verbatim}\\end{document}");
        ReaderChunk[] chunks = LatexReaderAdapter.Read(document, readerOptions: new ReaderOptions { MaxChars = 32 },
            latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks, IncludeDiagnostics = false }).ToArray();
        Assert.True(chunks.Length > 1);
        string expected = Assert.Single(document.ToMarkdownDocument().Blocks.OfType<OfficeIMO.Markdown.CodeBlock>()).Content;
        Assert.Equal(expected, string.Concat(chunks.Select(static chunk => chunk.Text)));
        Assert.All(chunks, static chunk => {
            Assert.InRange(chunk.Text.Length, 1, 32);
            OfficeIMO.Markdown.MarkdownDoc parsed = OfficeIMO.Markdown.MarkdownReader.Parse(chunk.Markdown!);
            Assert.All(parsed.Blocks, static block => Assert.IsType<OfficeIMO.Markdown.CodeBlock>(block));
            Assert.DoesNotContain("<a id=\"fake\"></a>", parsed.ToHtmlFragment(), StringComparison.Ordinal);
        });
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SplitLiteralProseDoesNotActivateMarkdownOrHtml(bool blocks) {
        string prose = string.Concat(Enumerable.Repeat("\\# Heading *value* [link](url) <a id=\"fake\"></a> ", 6));
        ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.Parse("\\documentclass{article}\\begin{document}" + prose + "\\end{document}"),
            readerOptions: new ReaderOptions { MaxChars = 32 }, latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks, IncludeDiagnostics = false }).ToArray();
        Assert.True(chunks.Length > 1);
        Assert.All(chunks, static chunk => {
            OfficeIMO.Markdown.MarkdownDoc parsed = OfficeIMO.Markdown.MarkdownReader.Parse(chunk.Markdown!);
            Assert.All(parsed.Blocks, static block => Assert.IsType<OfficeIMO.Markdown.ParagraphBlock>(block));
            Assert.DoesNotContain(parsed.Descendants(), static node => node is OfficeIMO.Markdown.LinkInline || node is OfficeIMO.Markdown.ItalicSequenceInline);
            Assert.DoesNotContain("<a id=\"fake\"></a>", parsed.ToHtmlFragment(), StringComparison.Ordinal);
        });
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SplitInlineCodeAndRealAnchorsRetainDistinctProvenance(bool blocks) {
        string body = new string('p', 70) + "\\label{real}\\verb|<a id=\"fake\"></a> *value*|" + new string('s', 70);
        ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.Parse("\\documentclass{article}\\begin{document}" + body + "\\end{document}"),
            readerOptions: new ReaderOptions { MaxChars = 32 }, latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks, IncludeDiagnostics = false }).ToArray();
        string html = string.Concat(chunks.Select(static chunk => OfficeIMO.Markdown.MarkdownReader.Parse(chunk.Markdown!).ToHtmlFragment()));
        Assert.Contains("<a id=\"real\"></a>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<a id=\"fake\"></a>", html, StringComparison.Ordinal);
        Assert.Contains(chunks, static chunk => OfficeIMO.Markdown.MarkdownReader.Parse(chunk.Markdown!).Blocks.OfType<OfficeIMO.Markdown.CodeBlock>().Any());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SplitMetadataCarriesGlobalWarningsOnce(bool blocks) {
        string source = "\\documentclass{article}\\author{" + new string('a', 100) + "\\begin{comment}PRIVATE\\end{comment}}\\begin{document}\\end{document}";
        ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.Parse(source), readerOptions: new ReaderOptions { MaxChars = 12 },
            latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks }).ToArray();
        Assert.True(chunks.Length > 1);
        Assert.Single(chunks.SelectMany(static chunk => chunk.Warnings ?? Array.Empty<string>()), static warning => warning.StartsWith("LATEXMD210:", StringComparison.Ordinal));
        Assert.All(chunks, static chunk => Assert.Contains(chunk.Warnings ?? Array.Empty<string>(), static warning => warning.Contains("MaxChars", StringComparison.Ordinal)));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TinySplitLimitsRetainCompleteUnicodeScalars(bool blocks) {
        const string text = "\U0001F600\U0001F642";
        ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.Parse("\\documentclass{article}\\begin{document}" + text + "\\end{document}"),
            readerOptions: new ReaderOptions { MaxChars = 1 }, latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks, IncludeDiagnostics = false }).ToArray();
        Assert.Equal(text, string.Concat(chunks.Select(static chunk => chunk.Text)));
        Assert.Equal(2, chunks.Length);
        Assert.All(chunks, static chunk => Assert.Equal(2, chunk.Text.Length));
    }

    [Theory]
    [InlineData(true, "    ")]
    [InlineData(false, "    ")]
    [InlineData(true, "\t")]
    [InlineData(false, " \t")]
    public void SplitIndentedProseRetainsLiteralTextInsteadOfBecomingCode(bool blocks, string indentation) {
        const string literal = "*value* <a id=\"fake\"></a>";
        string body = new string('p', 70) + "\n" + indentation + literal;
        ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.Parse("\\documentclass{article}\\begin{document}" + body + "\\end{document}"),
            readerOptions: new ReaderOptions { MaxChars = 40 }, latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks, IncludeDiagnostics = false }).ToArray();
        ReaderChunk part = Assert.Single(chunks, static chunk => chunk.Text.Contains("*value*", StringComparison.Ordinal));
        OfficeIMO.Markdown.ParagraphBlock paragraph = Assert.IsType<OfficeIMO.Markdown.ParagraphBlock>(
            Assert.Single(OfficeIMO.Markdown.MarkdownReader.Parse(part.Markdown!).Blocks));
        var text = new StringBuilder();
        foreach (OfficeIMO.Markdown.IPlainTextMarkdownInline inline in paragraph.Inlines.Nodes.OfType<OfficeIMO.Markdown.IPlainTextMarkdownInline>()) inline.AppendPlainText(text);
        Assert.Equal(indentation + literal, text.ToString());
    }

    [Fact]
    public void ParserWarningsRemainAttachedToTheirOwnBlocks() {
        LatexDocument document = LatexDocument.Parse("\\begin{document}First }\n\nSecond }\n\nLast\\end{document}");
        ReaderChunk[] chunks = LatexReaderAdapter.Read(document).ToArray();
        Assert.Equal(3, chunks.Length);
        Assert.All(chunks.Take(2), static chunk => Assert.Single(chunk.Warnings ?? Array.Empty<string>(), static warning => warning.StartsWith("LATEX002:", StringComparison.Ordinal)));
        Assert.DoesNotContain(chunks[2].Warnings ?? Array.Empty<string>(), static warning => warning.StartsWith("LATEX002:", StringComparison.Ordinal));
    }

    [Fact]
    public void SplitFirstBlockCarriesGlobalWarningsOnlyOnce() {
        LatexDocument document = LatexDocument.Parse("\\author\\begin{document}" + new string('x', 100) + "\\end{document}");
        ReaderChunk[] chunks = LatexReaderAdapter.Read(document, readerOptions: new ReaderOptions { MaxChars = 12 }).ToArray();
        Assert.True(chunks.Length > 1);
        Assert.Single(chunks.SelectMany(static chunk => chunk.Warnings ?? Array.Empty<string>()), static warning => warning.StartsWith("LATEX007:", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(true, true)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(false, false)]
    public void MarkdownOnlyAnchorsRemainVisibleWithOrWithoutDiagnostics(bool blocks, bool diagnostics) {
        LatexDocument document = LatexDocument.Parse("\\begin{document}\\label{x}\\end{document}");
        ReaderChunk chunk = Assert.Single(LatexReaderAdapter.Read(document, latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks, IncludeDiagnostics = diagnostics }));
        Assert.Contains("<a id=\"x\"></a>", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("<a id=\"x\"></a>", chunk.Markdown, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(true, true)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(false, false)]
    public void MetadataDiagnosticsReachOneChunkEvenWhenTheBodyIsEmpty(bool blocks, bool empty) {
        LatexDocument document = LatexDocument.Parse("\\author{Public\\begin{comment}PRIVATE\\end{comment}}\\begin{document}" +
            (empty ? "" : "Body one.\n\nBody two.") + "\\end{document}");
        ReaderChunk[] chunks = LatexReaderAdapter.Read(document, latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks }).ToArray();
        Assert.NotEmpty(chunks);
        Assert.Single(chunks.SelectMany(static chunk => chunk.Warnings ?? Array.Empty<string>()), static warning => warning.StartsWith("LATEXMD210:", StringComparison.Ordinal));
        Assert.All(chunks, static chunk => Assert.DoesNotContain("PRIVATE", chunk.Markdown, StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void FigureCaptionsAppearOncePerFigureWithMultipleImages(bool blocks) {
        const string figure = "\\begin{figure}\\includegraphics{a.png}\\includegraphics{b.png}\\caption{Shared caption}\\end{figure}";
        ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.Parse("\\begin{document}" + figure + figure + "\\end{document}"),
            latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks }).ToArray();
        string text = string.Join("\n", chunks.Select(static chunk => chunk.Text));
        Assert.Equal(2, text.Split(new[] { "Shared caption" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(2, text.Split(new[] { "a.png" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(2, text.Split(new[] { "b.png" }, StringSplitOptions.None).Length - 1);
    }

    [Theory]
    [InlineData(8, 100)]
    [InlineData(100, 8)]
    public void DirectNativeLoadingHonorsTheSmallerByteLimitAndRestoresSeekableStreamState(int readerLimit, int nativeLimit) {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(Source));
        stream.Position = 5;
        Assert.Throws<IOException>(() => LatexReaderAdapter.Read(stream, readerOptions: new ReaderOptions { MaxInputBytes = readerLimit },
            latexOptions: new ReaderLatexOptions { ParseOptions = new LatexParseOptions { MaximumInputBytes = nativeLimit } }).ToArray());
        Assert.Equal(5, stream.Position);
        Assert.True(stream.CanRead);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EmptyMalformedDocumentsStillCarryParserWarnings(bool blocks) {
        ReaderChunk chunk = Assert.Single(LatexReaderAdapter.Read(LatexDocument.Parse("\\begin{document}"),
            latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks }));
        Assert.Contains(chunk.Warnings ?? Array.Empty<string>(), static warning => warning.StartsWith("LATEX004:", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PreambleParserWarningsReachTheBodyChunk(bool blocks) {
        ReaderChunk chunk = Assert.Single(LatexReaderAdapter.Read(LatexDocument.Parse("\\author\\begin{document}Body\\end{document}"),
            latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks }));
        Assert.Contains(chunk.Warnings ?? Array.Empty<string>(), static warning => warning.StartsWith("LATEX007:", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MetadataOnlyChunksRespectTheCharacterLimit(bool blocks) {
        ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.Parse("\\author{" + new string('a', 100) + "}\\begin{document}\\end{document}"),
            readerOptions: new ReaderOptions { MaxChars = 12 }, latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks }).ToArray();
        Assert.True(chunks.Length > 1);
        Assert.All(chunks, static chunk => {
            Assert.InRange(chunk.Text.Length, 1, 12);
            Assert.False(string.IsNullOrEmpty(chunk.Markdown)); // escaping and fences may add carrier overhead
            Assert.Contains(chunk.Warnings ?? Array.Empty<string>(), static warning => warning.Contains("MaxChars", StringComparison.Ordinal));
        });
    }

    [Fact]
    public void NativeParserLimitExceptionsRetainTheirType() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(Source));
        Assert.Throws<InvalidDataException>(() => LatexReaderAdapter.Read(stream,
            latexOptions: new ReaderLatexOptions { ParseOptions = new LatexParseOptions { MaximumInputLength = 8 } }).ToArray());
        Assert.Equal(0, stream.Position);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ParserDiagnosticsInEmptyProjectedBlocksRemainVisible(bool blocks) {
        ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.Parse("\\begin{document}}\n\nBody\\end{document}"),
            latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks }).ToArray();
        Assert.Single(chunks.SelectMany(static chunk => chunk.Warnings ?? Array.Empty<string>()),
            static warning => warning.StartsWith("LATEX002:", StringComparison.Ordinal));
        Assert.Contains(chunks, static chunk => chunk.Text.Contains("Body", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(true, true)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(false, false)]
    public void AnchorsRemainVisibleBesideSplitTextAndInsideSplitParagraphs(bool blocks, bool inline) {
        string separator = inline ? string.Empty : "\n\n";
        string first = new string('p', 300);
        string second = new string('s', 300);
        foreach (string body in new[] { "\\label{x}" + separator + first,
            first + separator + "\\label{x}" + separator + second, first + separator + "\\label{x}",
            first + separator + "\\textbf{\\label{x}}" + separator + second }) {
            ReaderChunk[] chunks = LatexReaderAdapter.Read(LatexDocument.Parse("\\begin{document}" + body + "\\end{document}"),
                readerOptions: new ReaderOptions { MaxChars = 256 }, latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks, IncludeDiagnostics = false }).ToArray();
            Assert.Contains(chunks, static chunk => chunk.Markdown!.Contains("<a id=\"x\"></a>", StringComparison.Ordinal));
            Assert.All(chunks, static chunk => { Assert.InRange(chunk.Text.Length, 1, 256); Assert.InRange(chunk.Markdown!.Length, 1, 256); });
            string text = string.Concat(chunks.Select(static chunk => chunk.Text));
            Assert.Equal(300, text.Count(static character => character == 'p'));
            Assert.Equal(body.Contains(second, StringComparison.Ordinal) ? 300 : 0, text.Count(static character => character == 's'));
        }
    }

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

        Assert.Contains("MaxInputBytes", exception.Message, StringComparison.Ordinal);
        Assert.True(stream.CanRead);
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

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CustomListLabelsRemainInReaderTextMarkdownAndWarnings(bool blocks) {
        LatexDocument document = LatexDocument.Parse("\\begin{document}\\begin{itemize}\\item[URGENT] Call now\\end{itemize}\\end{document}");
        ReaderChunk chunk = Assert.Single(LatexReaderAdapter.Read(document, latexOptions: new ReaderLatexOptions { ChunkByBlock = blocks }));
        Assert.Contains("URGENT: Call now", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("URGENT: Call now", chunk.Markdown, StringComparison.Ordinal);
        Assert.Contains(chunk.Warnings ?? Array.Empty<string>(), static warning => warning.StartsWith("LATEXMD214:", StringComparison.Ordinal));
    }

    [Fact]
    public void InactiveDefinitionFallbackKeepsPublicSourceAndSuppressesNestedComments() {
        const string source = "\\begin{document}\\newcommand{\\draft}{Public\\begin{comment}PRIVATE DRAFT\\end{comment}}Actual\\end{document}";
        ReaderChunk chunk = Assert.Single(LatexReaderAdapter.Read(LatexDocument.Parse(source)));
        Assert.Contains("Public", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("Actual", chunk.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE", chunk.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE", chunk.Markdown, StringComparison.Ordinal);
        Assert.Contains(chunk.Warnings ?? Array.Empty<string>(), static warning => warning.StartsWith("LATEXMD210:", StringComparison.Ordinal));
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
        Assert.Throws<IOException>(() => LatexReaderAdapter.Read(limited, latexOptions: clone).ToArray());
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
