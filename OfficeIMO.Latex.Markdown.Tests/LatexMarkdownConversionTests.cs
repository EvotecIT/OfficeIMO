namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexMarkdownConversionTests {
    private const string Source =
        "\\documentclass{article}\n" +
        "\\title{Guide}\n\\author{Author}\n" +
        "\\begin{document}\n\\maketitle\n" +
        "\\section{Start}\\label{sec:start}\n" +
        "Text with \\textbf{bold}, \\ref{sec:start}, \\cite{key}, and $x^2$.\n\n" +
        "\\begin{itemize}\n\\item One\n\\item Two\n\\end{itemize}\n" +
        "\\begin{figure}\n\\includegraphics{plot.png}\n\\caption{Plot}\n\\label{fig:plot}\n\\end{figure}\n" +
        "\\begin{tabular}{ll}\nName & Value\\\\\nA & B\\\\\n\\end{tabular}\n" +
        "\\begin{theorem}[Result]Proof text.\\end{theorem}\n" +
        "\\begin{equation}E=mc^2\\end{equation}\n" +
        "\\end{document}\n";

    [Fact]
    public void TechnicalLatex_ConvertsToTypedMarkdownWithExplicitMathAndCitationLoss() {
        LatexToMarkdownResult result = LatexDocument.Parse(Source).Document.ToMarkdownDocumentResult();

        Assert.Equal("Guide", result.Value.Blocks.OfType<HeadingBlock>().First().Text);
        Assert.Contains(result.Value.Blocks.OfType<HeadingBlock>(), heading => heading.Text == "Start");
        ParagraphBlock paragraph = Assert.Single(result.Value.Blocks.OfType<ParagraphBlock>(), block =>
            block.Inlines.Nodes.OfType<BoldSequenceInline>().Any());
        Assert.Contains(paragraph.Inlines.Nodes, node => node is LinkInline);
        Assert.Contains(paragraph.Inlines.Nodes, node => node is CodeSpanInline);
        Assert.Single(result.Value.Blocks.OfType<UnorderedListBlock>());
        ImageBlock image = Assert.Single(result.Value.Blocks.OfType<ImageBlock>());
        Assert.Equal("plot.png", image.Path);
        Assert.Equal("Plot", image.Caption);
        Assert.Single(result.Value.Blocks.OfType<TableBlock>());
        Assert.Single(result.Value.Blocks.OfType<CalloutBlock>());
        Assert.Single(result.Value.Blocks.OfType<SemanticFencedBlock>());
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD101");
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD102");
        Assert.Equal(Source, LatexDocument.Parse(Source).Document.ToLatex());
    }

    [Fact]
    public void LatexProjection_UsesExistingWordBridge() {
        LatexToMarkdownResult conversion = LatexDocument.Parse(Source).Document.ToMarkdownDocumentResult();

        using var word = conversion.Value.ToWordDocument();
        string text = string.Join(" ", word.Paragraphs.Select(paragraph => paragraph.Text));

        Assert.Contains("Guide", text, StringComparison.Ordinal);
        Assert.Contains("Start", text, StringComparison.Ordinal);
        Assert.Contains("bold", text, StringComparison.Ordinal);
        Assert.True(word.Tables.Count > 0);
    }

    [Fact]
    public void DirectPdfAdapter_PreservesProjectionDiagnosticsAndProducesPdf() {
        var result = LatexDocument.Parse(Source).Document.ToPdfDocumentResult();
        byte[] bytes = result.ToBytes();

        Assert.True(bytes.Length > 100);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(bytes, 0, 5));
        Assert.Contains(result.Warnings, warning =>
            warning.Converter == "OfficeIMO.Latex.Pdf" &&
            warning.Code == "LATEXMD101" &&
            warning.Details["stage"] == "semantic-projection");
        Assert.Contains(result.Warnings, warning => warning.Code == "LATEXMD102");
    }

    [Fact]
    public void RepresentativeMarkdown_GeneratesRecognizedLosslessLatexProfile() {
        const string markdown =
            "---\ntitle: Guide\nauthor: Author\n---\n\n" +
            "# Guide\n\n## Start\n\nParagraph with **bold** and [section](#start).\n\n" +
            "- One\n- Two\n\n" +
            "| Name | Value |\n| --- | --- |\n| A | B |\n\n" +
            "![Plot](plot.png)\n";
        MarkdownDoc document = MarkdownReader.Parse(markdown);

        MarkdownToLatexResult result = document.ToLatexDocumentResult();

        Assert.Contains("\\documentclass{article}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\usepackage{graphicx}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\usepackage{hyperref}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\title{Guide}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\section{Start}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\begin{itemize}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\begin{tabular}{ll}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\begin{figure}", result.Source, StringComparison.Ordinal);
        Assert.True(result.Value.IsRecognizedProfile);
        Assert.Equal(result.Source, result.Value.ToLatex());
        Assert.Single(result.Value.Lists);
        Assert.Single(result.Value.Tables);
        Assert.Single(result.Value.Figures);
    }

    [Fact]
    public void StructuredMarkdownTableSpans_GenerateMulticolumnAndMultirow() {
        const string asciidocLikeMarkdown = "| H1 | H2 |\n| --- | --- |\n| A | B |\n";
        TableBlock table = Assert.Single(MarkdownReader.Parse(asciidocLikeMarkdown).Blocks.OfType<TableBlock>());
        TableCell cell = table.GetCell(0, 0)!;
        cell.ColumnSpan = 2;
        cell.RowSpan = 2;
        MarkdownDoc document = MarkdownDoc.Create().Add(table);

        MarkdownToLatexResult result = document.ToLatexDocumentResult();

        Assert.Contains("\\usepackage{multirow}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\multicolumn{2}{l}{", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\multirow{2}{*}{", result.Source, StringComparison.Ordinal);
    }

    [Fact]
    public void UnsupportedMarkdown_IsVisibleInVerbatimAndDiagnosed() {
        MarkdownDoc document = MarkdownDoc.Create().Hr();

        MarkdownToLatexResult result = document.ToLatexDocumentResult();

        Assert.Contains("\\begin{verbatim}", result.Source, StringComparison.Ordinal);
        Assert.Equal(LatexMarkdownConversionOutcome.SourceFallback, Assert.Single(result.Report.Diagnostics).Outcome);
    }

    [Fact]
    public void MarkdownFootnoteReference_UsesDefinitionBodyAndDoesNotEmitDefinitionFallback() {
        MarkdownDoc document = MarkdownReader.Parse("Text[^1].\n\n[^1]: real **text**\n");

        MarkdownToLatexResult result = document.ToLatexDocumentResult();

        Assert.Contains("\\footnote{real \\textbf{text}}", result.Source, StringComparison.Ordinal);
        Assert.DoesNotContain("\\footnote{1}", result.Source, StringComparison.Ordinal);
        Assert.DoesNotContain("\\begin{verbatim}", result.Source, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Feature == nameof(FootnoteDefinitionBlock));
    }
}
