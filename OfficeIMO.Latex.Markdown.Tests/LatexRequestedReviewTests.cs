namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexRequestedReviewTests {
    [Fact]
    public void StrikePackageDoesNotReplaceOrdinaryEmphasis() {
        var result = MarkdownReader.Parse("*Italic marker* and ~~Removed marker~~").ToLatexDocumentResult();
        Assert.Contains(result.Value.Commands, command => command.Name == "usepackage" &&
            command.GetRequiredArgument(0)?.Content == "ulem" && command.GetOptionalArgument(0)?.Content == "normalem");
        Assert.Contains("*Italic marker*", result.Value.ToMarkdownDocument().ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains("~~Removed marker~~", result.Value.ToMarkdownDocument().ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void NonSequentialAuthoredMarkersAreDiagnosed() {
        var result = MarkdownReader.Parse("1. first\n3. third\n").ToLatexDocumentResult();
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "MDLATEX022");
        Assert.Contains("third", result.Source, StringComparison.Ordinal);
    }

    [Fact]
    public void LinkMetadataAndHighlightSimplificationAreReportedWithoutLosingVisibleContent() {
        var inlines = new InlineSequence().Link("Visible", "https://example.test", "tooltip", "_blank", "noopener").Highlight("Highlighted");
        var result = MarkdownDoc.Create().Add(new ParagraphBlock(inlines)).ToLatexDocumentResult();
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "MDLATEX027");
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "MDLATEX028");
        Assert.Contains("[Visible](https://example.test)", result.Value.ToMarkdownDocument().ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains("**Highlighted**", result.Value.ToMarkdownDocument().ToMarkdown(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(@"before\end{verbatim}\section{Literal}after")]
    [InlineData(@"before\end {verbatim}\section{Literal}after")]
    [InlineData("before\\end% comment\n{verbatim}\\section{Literal}after")]
    [InlineData(@"before\end{ verbatim }\section{Literal}after")]
    public void RecognizedVerbatimClosingFormsCannotActivateLiteralCode(string content) {
        var result = MarkdownDoc.Create().Add(new CodeBlock("text", content)).ToLatexDocumentResult();
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "MDLATEX021");
        Assert.Empty(result.Value.Headings);
        Assert.DoesNotContain(result.Value.Diagnostics, diagnostic => diagnostic.Code == "LATEX005");
        var projected = result.Value.ToMarkdownDocument();
        Assert.Single(projected.Blocks.OfType<CodeBlock>());
        Assert.Contains(@"\section{Literal}", Assert.Single(projected.Blocks.OfType<CodeBlock>()).Content, StringComparison.Ordinal);
    }

    [Fact]
    public void PromotedTitleAndFragmentLinksRetainFormattingLabelsAndVisibleText() {
        var markdown = MarkdownReader.Parse("# A **rich** title {#intro}\n\n[Read **title**](#intro)", new MarkdownReaderOptions { GenericAttributes = true });
        var result = markdown.ToLatexDocumentResult();
        Assert.False(result.Report.HasLoss);
        Assert.Contains(@"\title{A \textbf{rich} title}", result.Source, StringComparison.Ordinal);
        Assert.Contains(result.Value.Labels, label => label.Name == "intro");
        Assert.Contains("[Read **title**](#intro)", result.Value.ToMarkdownDocument().ToMarkdown(), StringComparison.Ordinal);
        Assert.NotEmpty(result.Value.ToPdfBytes());
    }

    [Fact]
    public void LooseListParagraphsRemainInOrder() {
        var markdown = MarkdownReader.Parse("- First paragraph\n\n  IMPORTANT SECOND PARAGRAPH\n\n  - Nested item\n");
        string source = markdown.ToLatexDocumentResult().Source;
        int first = source.IndexOf("First paragraph", StringComparison.Ordinal);
        int second = source.IndexOf("IMPORTANT SECOND PARAGRAPH", StringComparison.Ordinal);
        int nested = source.IndexOf("Nested item", StringComparison.Ordinal);
        Assert.True(first >= 0 && second > first && nested > second);
        Assert.Contains("IMPORTANT SECOND PARAGRAPH", LatexDocument.Parse(source).ToMarkdownDocument().ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void ReverseLossReportsCoverCustomNumberingTaskStateCalloutAndImageMetadata() {
        var document = MarkdownReader.Parse("5. first\n6. second\n\n- [x] done\n- [ ] pending\n");
        document.Callout("warning", "Retained title", "Body");
        document.Add(new ImageBlock("figure.png", "alternate", "title"));
        var result = document.ToLatexDocumentResult();
        Assert.True(result.Report.HasLoss);
        foreach (string code in new[] { "MDLATEX022", "MDLATEX023", "MDLATEX025", "MDLATEX026" })
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == code);
        Assert.Contains("Retained title", result.Source, StringComparison.Ordinal);
        Assert.Contains(@"\texttt{[x]}", result.Source, StringComparison.Ordinal);
        Assert.Contains(@"\texttt{[ ]}", result.Source, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(@"\begin{theorem}\label{thm:x}Body\end{theorem} See \ref{thm:x}", "Body")]
    [InlineData(@"\begin{table}\label{tab:x}\begin{tabular}{l}Cell\\\end{tabular}\end{table} See \ref{tab:x}", "Cell")]
    [InlineData(@"\begin{figure}\label{fig:x}\caption{Figure marker}\end{figure} See \ref{fig:x}", "Figure marker")]
    [InlineData(@"\label{p:x}Body. See \ref{p:x}", "Body")]
    public void NonHeadingLabelsProduceWorkingPdfDestinations(string body, string marker) {
        var document = LatexDocument.Parse(@"\documentclass{article}\begin{document}" + body + @"\end{document}");
        byte[] bytes = document.ToPdfBytes();
        Assert.Contains(marker, OfficeIMO.Pdf.PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void MissingReferencesRetainVisibleTextWithAnExportDiagnostic() {
        var document = LatexDocument.Parse(@"\documentclass{article}\begin{document}See \ref{missing:x}\end{document}");
        var result = document.ToPdfDocumentResult();
        byte[] bytes = result.Value.ToBytes();
        Assert.Contains("missing:x", OfficeIMO.Pdf.PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(result.Warnings, warning => warning.Code == "UnresolvedInternalLink");
    }

    [Theory]
    [InlineData(@"TRAILING\section{Inactive}")]
    [InlineData(@"TRAILING\[inactive\]")]
    public void ProjectionDoesNotExposeInertSourceAfterDocumentEnd(string trailer) {
        var document = LatexDocument.Parse(@"\documentclass{article}\begin{document}Public\end{document}" + trailer);
        Assert.Equal("Public", Assert.IsType<MarkdownTextRun>(Assert.Single(Assert.Single(document.ToMarkdownDocument().Blocks.OfType<ParagraphBlock>()).Inlines.Nodes)).Text);
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(document.ToPdfBytes()).ExtractText();
        Assert.Contains("Public", text, StringComparison.Ordinal);
        Assert.DoesNotContain("TRAILING", text, StringComparison.Ordinal);
    }
}
