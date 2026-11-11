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
        var result = LatexDocument.ParseResult(Source).Document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport1);
        var resultReport = (LatexToMarkdownReport)conversionReport1;

        Assert.Equal("Guide", result.Blocks.OfType<HeadingBlock>().First().Text);
        Assert.Contains(result.Blocks.OfType<HeadingBlock>(), heading => heading.Text == "Start");
        ParagraphBlock paragraph = Assert.Single(result.Blocks.OfType<ParagraphBlock>(), block =>
            block.Inlines.Nodes.OfType<BoldSequenceInline>().Any());
        Assert.Contains(paragraph.Inlines.Nodes, node => node is LinkInline);
        Assert.Equal("x^2", Assert.Single(paragraph.Inlines.Nodes.OfType<MathInline>()).Source);
        Assert.Single(result.Blocks.OfType<UnorderedListBlock>());
        ImageBlock image = Assert.Single(result.Blocks.OfType<ImageBlock>());
        Assert.Equal("plot.png", image.Path);
        Assert.Equal("Plot", image.Caption);
        Assert.Single(result.Blocks.OfType<TableBlock>());
        Assert.Single(result.Blocks.OfType<CalloutBlock>());
        Assert.Single(result.Blocks.OfType<SemanticFencedBlock>());
        Assert.Contains(resultReport.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD102");
        Assert.Equal(Source, LatexDocument.ParseResult(Source).Document.ToLatex());
    }

    [Fact]
    public void LatexProjection_UsesExistingWordBridge() {
        var conversion = LatexDocument.ParseResult(Source).Document.ToMarkdown(null);

        using var word = conversion.ToWord();
        string text = string.Join(" ", word.Paragraphs.Select(paragraph => paragraph.Text));

        Assert.Contains("Guide", text, StringComparison.Ordinal);
        Assert.Contains("Start", text, StringComparison.Ordinal);
        Assert.Contains("bold", text, StringComparison.Ordinal);
        Assert.True(word.Tables.Count > 0);
    }

    [Fact]
    public void DirectPdfAdapter_PreservesProjectionDiagnosticsAndProducesPdf() {
        var result = LatexDocument.ParseResult(Source).Document.ConvertToPdfWithEvidence(null, default);
        byte[] bytes = result.ToBytes();
        LatexToMarkdownReport projection = Assert.IsType<LatexToMarkdownReport>(
            Assert.Single(result.SourceConversionReports));

        Assert.True(bytes.Length > 100);
        Assert.Equal("%PDF-", Encoding.ASCII.GetString(bytes, 0, 5));
        Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD102");
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "LATEXMD102");
        Assert.True(result.HasLoss);
    }

    [Fact]
    public void RepresentativeMarkdown_GeneratesRecognizedLosslessLatexProfile() {
        const string markdown =
            "---\ntitle: Guide\nauthor: Author\n---\n\n" +
            "# Guide\n\n## Start\n\nParagraph with **bold** and [section](#start).\n\n" +
            "- One\n- Two\n\n" +
            "| Name | Value |\n| --- | --- |\n| A | B |\n\n" +
            "![Plot](plot.png)\n";
        MarkdownDocument document = MarkdownDocument.Parse(markdown);

        var result = document.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport3);
        var resultReport = (MarkdownToLatexReport)conversionReport3;

        Assert.Contains("\\documentclass{article}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\usepackage{graphicx}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\usepackage{hyperref}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\title{Guide}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\section{Start}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\begin{itemize}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\begin{tabular}{ll}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\begin{figure}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.True(result.IsRecognizedProfile);
        Assert.Equal(resultReport.GeneratedSource, result.ToLatex());
        Assert.Single(result.Lists);
        Assert.Single(result.Tables);
        Assert.Single(result.Figures);
    }

    [Fact]
    public void StructuredMarkdownTableSpans_GenerateMulticolumnAndMultirow() {
        const string asciidocLikeMarkdown = "| H1 | H2 |\n| --- | --- |\n| A | B |\n";
        TableBlock table = Assert.Single(MarkdownDocument.Parse(asciidocLikeMarkdown).Blocks.OfType<TableBlock>());
        TableCell cell = table.GetCell(0, 0)!;
        cell.ColumnSpan = 2;
        cell.RowSpan = 2;
        MarkdownDocument document = MarkdownDocument.Create();
        document.Add(table);

        var result = document.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport4);
        var resultReport = (MarkdownToLatexReport)conversionReport4;

        Assert.Contains("\\usepackage{multirow}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\multicolumn{2}{l}{", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\multirow{2}{*}{", resultReport.GeneratedSource, StringComparison.Ordinal);
    }

    [Fact]
    public void UnsupportedMarkdown_IsVisibleInVerbatimAndDiagnosed() {
        MarkdownDocument document = MarkdownDocument.Create().Hr();

        var result = document.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport5);
        var resultReport = (MarkdownToLatexReport)conversionReport5;

        Assert.Contains("\\begin{verbatim}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Equal(LatexMarkdownConversionOutcome.SourceFallback, Assert.Single(resultReport.NativeDiagnostics).Outcome);
    }

    [Fact]
    public void MarkdownFootnoteReference_UsesDefinitionBodyAndDoesNotEmitDefinitionFallback() {
        MarkdownDocument document = MarkdownDocument.Parse("Text[^1].\n\n[^1]: real **text**\n");

        var result = document.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport6);
        var resultReport = (MarkdownToLatexReport)conversionReport6;

        Assert.Contains("\\footnote{real \\textbf{text}}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.DoesNotContain("\\footnote{1}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.DoesNotContain("\\begin{verbatim}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.DoesNotContain(resultReport.NativeDiagnostics, diagnostic => diagnostic.Feature == nameof(FootnoteDefinitionBlock));
    }
}
