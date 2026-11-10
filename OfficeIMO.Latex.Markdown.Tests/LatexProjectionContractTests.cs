namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexProjectionContractTests {
    private static string Wrap(string body) => "\\documentclass{article}\n\\begin{document}\n" + body + "\n\\end{document}";

    [Theory]
    [InlineData("A\\textasciitilde B", "A~B")]
    [InlineData("A\\textbackslash B", "A\\B")]
    [InlineData("A\\textasciicircum B", "A^B")]
    [InlineData("A\\textasciitilde \t\r\n B", "A~B")]
    [InlineData("A\\textasciitilde% comment\n  B", "A~B")]
    [InlineData("A\\textasciitilde{} B", "A~ B")]
    [InlineData("A\\textasciitilde \t{} B", "A~ B")]
    [InlineData("A\\textasciitilde{B} C", "A~B C")]
    [InlineData("A\\% B", "A% B")]
    public void LiteralControlWordsDiscardDelimiterWhitespaceInTextAndScalars(string source, string expected) {
        MarkdownDocument markdown = LatexDocument.Parse("\\author{" + source + "}" + Wrap(source)).ToMarkdown();
        Assert.Equal(expected, markdown.FindFrontMatterEntry("author")?.Value.AsString());
        var text = new StringBuilder();
        foreach (IPlainTextMarkdownInline inline in Assert.Single(markdown.Blocks.OfType<ParagraphBlock>()).Inlines.Nodes.OfType<IPlainTextMarkdownInline>()) inline.AppendPlainText(text);
        Assert.Equal(expected, text.ToString());
    }

    [Fact]
    public void LiteralControlWordDelimitersDoNotConsumeParagraphsOrExplicitGroupSpaces() {
        MarkdownDocument markdown = LatexDocument.Parse(Wrap("A\\textasciitilde \n\nB\n\nC\\newline D")).ToMarkdown();
        ParagraphBlock[] paragraphs = markdown.Blocks.OfType<ParagraphBlock>().ToArray();
        Assert.Equal(3, paragraphs.Length);
        Assert.Equal("A~", PlainText(paragraphs[0].Inlines));
        Assert.Equal("B", PlainText(paragraphs[1].Inlines));
        Assert.Equal("C D", PlainText(paragraphs[2].Inlines));
        Assert.Single(paragraphs[2].Inlines.Nodes.OfType<HardBreakInline>());
    }

    [Fact]
    public void ResourceLiteralCommandsShareControlWordDelimiterHandling() {
        MarkdownDocument markdown = LatexDocument.Parse(Wrap("\\url{https://example.test/\\textasciitilde alice} " +
            "\\href{https://example.test/\\textasciitilde alice}{Link} \\includegraphics{\\textasciitilde images/a.png}")).ToMarkdown();
        InlineSequence inlines = Assert.Single(markdown.Blocks.OfType<ParagraphBlock>()).Inlines;
        Assert.All(inlines.Nodes.OfType<LinkInline>(), static link => Assert.Equal("https://example.test/~alice", link.Url));
        Assert.Equal("~images/a.png", Assert.Single(inlines.Nodes.OfType<ImageInline>()).Src);
    }

    [Theory]
    [InlineData("[]")]
    [InlineData("[ ]")]
    [InlineData("[\\label{item}]")]
    public void EmptyCustomListLabelsDoNotInjectVisiblePunctuation(string label) {
        MarkdownDocument markdown = LatexDocument.Parse(Wrap("\\begin{itemize}\\item" + label + " Body\\end{itemize}")).ToMarkdown();
        ListItem item = Assert.Single(Assert.Single(markdown.Blocks.OfType<UnorderedListBlock>()).Items);
        using var addedDocument = MarkdownDocument.Create();
        addedDocument.Add(new ParagraphBlock(item.Content));
        string rendered = addedDocument.ToMarkdown();
        Assert.DoesNotContain(":", rendered, StringComparison.Ordinal);
        Assert.Contains("Body", rendered, StringComparison.Ordinal);
        if (label.Contains("label")) Assert.Contains("<a id=\"item\"></a>", rendered, StringComparison.Ordinal);
    }

    private static string PlainText(InlineSequence inlines) {
        var text = new StringBuilder();
        foreach (IPlainTextMarkdownInline inline in inlines.Nodes.OfType<IPlainTextMarkdownInline>()) inline.AppendPlainText(text);
        return text.ToString();
    }

    [Fact]
    public void DisplayScalarsDecodeEscapesAndReportFlattenedFormatting() {
        string body = "\\begin{figure}\\includegraphics{~/a.png}\\caption{\\textbf{R\\&D} 100\\%}\\end{figure}\n" +
            "\\begin{table}\\caption{\\emph{Cost} 50\\%}\\begin{tabular}{l}A\\\\\\end{tabular}\\end{table}\n" +
            "\\begin{theorem}[\\textbf{Useful}~100\\%]Body\\end{theorem}";
        var result = LatexDocument.Parse("\\title{Cost 100\\%}\\author{\\textbf{Alice\\_Doe}}\\date{R\\&D}" + Wrap(body)).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport1);
        var resultReport = (LatexToMarkdownReport)conversionReport1;
        Assert.Equal("Cost 100%", result.FindFrontMatterEntry("title")?.Value.AsString());
        Assert.Equal("Alice_Doe", result.FindFrontMatterEntry("author")?.Value.AsString());
        Assert.Equal("R&D", result.FindFrontMatterEntry("date")?.Value.AsString());
        Assert.Equal("R&D 100%", Assert.Single(result.Blocks.OfType<ImageBlock>()).Caption);
        Assert.Equal("Cost 50%", Assert.Single(result.Blocks.OfType<TableBlock>()).Attributes.GetAttribute("caption"));
        Assert.Equal("Useful 100%", Assert.Single(result.Blocks.OfType<CalloutBlock>()).Title);
        foreach (string feature in new[] { "metadata:author", "figure-caption", "table-caption", "theorem-title" })
            Assert.Contains(resultReport.NativeDiagnostics, diagnostic => diagnostic.Code == "LATEXMD115" && diagnostic.Feature == feature);
    }

    [Fact]
    public void PlainMetadataAndFigureCaptionsRoundTripThroughTheWriter() {
        var source = MarkdownDocument.Create();
        source.FrontMatter(new Dictionary<string, object?> { ["author"] = "Alice_Doe", ["date"] = "R&D" });
        source.Add(new ImageBlock("a.png", "100% useful") { Caption = "100% useful" });
        MarkdownDocument result = source.ToLatex().ToMarkdown();
        Assert.Equal("Alice_Doe", result.FindFrontMatterEntry("author")?.Value.AsString());
        Assert.Equal("R&D", result.FindFrontMatterEntry("date")?.Value.AsString());
        Assert.Equal("100% useful", Assert.Single(result.Blocks.OfType<ImageBlock>()).Caption);
    }

    [Theory]
    [InlineData("\\textbf{A\\label{x}B}", "**AB**")]
    [InlineData("\\emph{A\\textbf{B\\label{x}C}D}", "**BC**")]
    [InlineData("$A\\label{x}B$", "`AB`")]
    public void TheoremLabelsDoNotSliceEnclosingInlineSyntax(string body, string expected) {
        var result = LatexDocument.Parse(Wrap("\\begin{theorem}" + body + "\\end{theorem}")).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport2);
        var resultReport = (LatexToMarkdownReport)conversionReport2;
        CalloutBlock callout = Assert.Single(result.Blocks.OfType<CalloutBlock>());
        Assert.Equal("x", callout.Attributes.ElementId);
        Assert.Contains(expected, ((result.ToMarkdown())), StringComparison.Ordinal);
        Assert.DoesNotContain("\\label", ((result.ToMarkdown())), StringComparison.Ordinal);
        if (body.StartsWith("\\textbf", StringComparison.Ordinal)) Assert.Empty(resultReport.Diagnostics);
    }

    [Fact]
    public void BatchedBlocksAndFigureImagesHaveCompleteObjectTreeNavigation() {
        MarkdownDocument result = LatexDocument.Parse(Wrap("First\n\n\\begin{figure}\\includegraphics{a.png}\\includegraphics{b.png}\\end{figure}\nLast")).ToMarkdown();
        Assert.Equal(4, result.Blocks.Count);
        for (int index = 0; index < result.Blocks.Count; index++) {
            MarkdownObject block = (MarkdownObject)result.Blocks[index];
            Assert.Same(result, block.Parent);
            Assert.Equal(index + 1, block.IndexInParent); // front matter occupies the first object slot
            Assert.Same(index == 0 ? result.DocumentHeader : result.Blocks[index - 1], block.PreviousSibling);
            Assert.Same(index + 1 == result.Blocks.Count ? null : result.Blocks[index + 1], block.NextSibling);
            foreach (MarkdownObject child in block.Descendants()) Assert.Same(result, child.Document);
        }
    }

    [Theory]
    [InlineData("~")]
    [InlineData("\\textasciitilde{}")]
    public void UrlAndHrefDestinationsRetainLiteralTildes(string tilde) {
        string destination = "https://example.test/" + tilde + "alice";
        MarkdownDocument markdown = LatexDocument.Parse(Wrap("\\url{" + destination + "} \\href{" + destination + "}{Alice}")).ToMarkdown();
        LinkInline[] links = Assert.Single(markdown.Blocks.OfType<ParagraphBlock>()).Inlines.Nodes.OfType<LinkInline>().ToArray();
        Assert.Equal(2, links.Length);
        Assert.All(links, static link => Assert.Equal("https://example.test/~alice", link.Url));
        Assert.Contains("https://example.test/~alice", ((markdown.ToMarkdown())), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InlineAndFigureResourcesRetainTildePaths(bool figure) {
        string image = "\\includegraphics{~/images/a.png}";
        MarkdownDocument markdown = LatexDocument.Parse(Wrap(figure ? "\\begin{figure}" + image + "\\end{figure}" : image)).ToMarkdown();
        string path = figure ? Assert.Single(markdown.Blocks.OfType<ImageBlock>()).Path
            : Assert.Single(Assert.Single(markdown.Blocks.OfType<ParagraphBlock>()).Inlines.Nodes.OfType<ImageInline>()).Src;
        Assert.Equal("~/images/a.png", path);
    }

    [Fact]
    public void ProseStillMapsTildeToSpacingAndExplicitTildeToLiteralText() {
        MarkdownDocument markdown = LatexDocument.Parse(Wrap("Hello~world \\textasciitilde{}")).ToMarkdown();
        var text = new StringBuilder();
        foreach (IPlainTextMarkdownInline inline in Assert.Single(markdown.Blocks.OfType<ParagraphBlock>()).Inlines.Nodes.OfType<IPlainTextMarkdownInline>()) inline.AppendPlainText(text);
        Assert.Equal("Hello world ~", text.ToString());
    }

    [Fact]
    public void IndexedProjectionKeepsCommentsLocalAndOpaquePercentVisible() {
        string body = "\\begin{unknown}First% PRIVATE1\n\\begin{comment}PRIVATE2\\end{comment}Tail\\end{unknown}\n" +
            "Next \\textbf{bold} \\verb|100%|% PRIVATE3\n\n" +
            "\\begin{unknown}Last% PRIVATE4\nEnd\\end{unknown}";
        var result = LatexDocument.Parse(Wrap(body)).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport3);
        var resultReport = (LatexToMarkdownReport)conversionReport3;
        Assert.Equal(3, result.Blocks.Count);
        string markdown = ((result.ToMarkdown()));
        Assert.DoesNotContain("PRIVATE", markdown, StringComparison.Ordinal);
        Assert.Contains("First", markdown, StringComparison.Ordinal);
        Assert.Contains("Tail", markdown, StringComparison.Ordinal);
        Assert.Contains("Last", markdown, StringComparison.Ordinal);
        Assert.Contains("100%", markdown, StringComparison.Ordinal);
        Assert.Single(resultReport.Diagnostics, static diagnostic => diagnostic.Code == "LATEXMD210");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RangeEnumerationHonorsCancellationAfterItHasStarted(bool comments) {
        LatexDocument document = LatexDocument.Parse(Wrap("\\textbf{a}% one\n\\textit{b}% two\n"));
        using var cancellation = new CancellationTokenSource();
        var context = new LatexProjectionContext(document, cancellation.Token);
        IEnumerable<object> values = comments ? context.Comments(document.SyntaxTree.Root.Span).Cast<object>()
            : context.InlineCandidates(document.Body!.ContentSpan.Start.Offset, document.Body.ContentSpan.End.Offset).Cast<object>();
        using IEnumerator<object> iterator = values.GetEnumerator();
        Assert.True(iterator.MoveNext());
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => iterator.MoveNext());
    }

    [Fact]
    public void SourceExtractionHonorsCancellationDuringInputEnumeration() {
        LatexDocument document = LatexDocument.Parse(Wrap("Visible text"));
        using var cancellation = new CancellationTokenSource();
        var context = new LatexProjectionContext(document, cancellation.Token);
        IEnumerable<LatexSourceSpan> Represented() {
            yield return document.Source.CreateSpan(0, 1);
            cancellation.Cancel();
            yield return document.Source.CreateSpan(1, 2);
        }
        Assert.Throws<OperationCanceledException>(() => context.ExtractResidual(document.SyntaxTree.Root.Span, Represented()));
    }
}
