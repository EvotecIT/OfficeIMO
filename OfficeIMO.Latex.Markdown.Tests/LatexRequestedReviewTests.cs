namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexRequestedReviewTests {
    [Fact]
    public void StrikePackageDoesNotReplaceOrdinaryEmphasis() {
        var result = MarkdownDocument.Parse("*Italic marker* and ~~Removed marker~~").ToLatex(null);
        Assert.Contains(result.Commands, command => command.Name == "usepackage" &&
            command.GetRequiredArgument(0)?.Content == "ulem" && command.GetOptionalArgument(0)?.Content == "normalem");
        Assert.Contains("*Italic marker*", ((result.ToMarkdown().ToMarkdown())), StringComparison.Ordinal);
        Assert.Contains("~~Removed marker~~", ((result.ToMarkdown().ToMarkdown())), StringComparison.Ordinal);
    }

    [Fact]
    public void NonSequentialAuthoredMarkersAreDiagnosed() {
        var result = MarkdownDocument.Parse("1. first\n3. third\n").ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport2);
        var resultReport = (MarkdownToLatexReport)conversionReport2;
        Assert.Contains(resultReport.Diagnostics, diagnostic => diagnostic.Code == "MDLATEX022");
        Assert.Contains("third", resultReport.GeneratedSource, StringComparison.Ordinal);
    }

    [Fact]
    public void LinkMetadataAndHighlightSimplificationAreReportedWithoutLosingVisibleContent() {
        var inlines = new InlineSequence().Link("Visible", "https://example.test", "tooltip", "_blank", "noopener").Highlight("Highlighted");
        using var addedDocument = MarkdownDocument.Create();
        addedDocument.Add(new ParagraphBlock(inlines));
        var result = addedDocument.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport3);
        var resultReport = (MarkdownToLatexReport)conversionReport3;
        Assert.Contains(resultReport.Diagnostics, diagnostic => diagnostic.Code == "MDLATEX027");
        Assert.Contains(resultReport.Diagnostics, diagnostic => diagnostic.Code == "MDLATEX028");
        Assert.Contains("[Visible](https://example.test)", ((result.ToMarkdown().ToMarkdown())), StringComparison.Ordinal);
        Assert.Contains("**Highlighted**", ((result.ToMarkdown().ToMarkdown())), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(@"before\end{verbatim}\section{Literal}after")]
    [InlineData(@"before\end {verbatim}\section{Literal}after")]
    [InlineData("before\\end% comment\n{verbatim}\\section{Literal}after")]
    [InlineData(@"before\end{ verbatim }\section{Literal}after")]
    public void RecognizedVerbatimClosingFormsCannotActivateLiteralCode(string content) {
        using var addedDocument = MarkdownDocument.Create();
        addedDocument.Add(new CodeBlock("text", content));
        var result = addedDocument.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport4);
        var resultReport = (MarkdownToLatexReport)conversionReport4;
        Assert.Contains(resultReport.Diagnostics, diagnostic => diagnostic.Code == "MDLATEX021");
        Assert.Empty(result.Headings);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "LATEX005");
        var projected = result.ToMarkdown();
        Assert.Single(projected.Blocks.OfType<CodeBlock>());
        Assert.Contains(@"\section{Literal}", Assert.Single(projected.Blocks.OfType<CodeBlock>()).Content, StringComparison.Ordinal);
    }

    [Fact]
    public void PromotedTitleAndFragmentLinksRetainFormattingLabelsAndVisibleText() {
        var markdown = MarkdownDocument.Parse("# A **rich** title {#intro}\n\n[Read **title**](#intro)", new MarkdownLoadOptions { GenericAttributes = true });
        var result = markdown.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport5);
        var resultReport = (MarkdownToLatexReport)conversionReport5;
        Assert.False(resultReport.HasLoss);
        Assert.Contains(@"\title{A \textbf{rich} title}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains(result.Labels, label => label.Name == "intro");
        Assert.Contains("[Read **title**](#intro)", ((result.ToMarkdown().ToMarkdown())), StringComparison.Ordinal);
        Assert.NotEmpty(result.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
    }

    [Fact]
    public void LooseListParagraphsRemainInOrder() {
        var markdown = MarkdownDocument.Parse("- First paragraph\n\n  IMPORTANT SECOND PARAGRAPH\n\n  - Nested item\n");
        markdown.ToLatex(null, out OfficeConversionReport report);
        string source = ((MarkdownToLatexReport)report).GeneratedSource;
        int first = source.IndexOf("First paragraph", StringComparison.Ordinal);
        int second = source.IndexOf("IMPORTANT SECOND PARAGRAPH", StringComparison.Ordinal);
        int nested = source.IndexOf("Nested item", StringComparison.Ordinal);
        Assert.True(first >= 0 && second > first && nested > second);
        Assert.Contains("IMPORTANT SECOND PARAGRAPH", ((LatexDocument.Parse(source).ToMarkdown().ToMarkdown())), StringComparison.Ordinal);
    }

    [Fact]
    public void ReverseLossReportsCoverCustomNumberingTaskStateCalloutAndImageMetadata() {
        var document = MarkdownDocument.Parse("5. first\n6. second\n\n- [x] done\n- [ ] pending\n");
        document.Callout("warning", "Retained title", "Body");
        document.Add(new ImageBlock("figure.png", "alternate", "title"));
        var result = document.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport7);
        var resultReport = (MarkdownToLatexReport)conversionReport7;
        Assert.True(resultReport.HasLoss);
        foreach (string code in new[] { "MDLATEX022", "MDLATEX023", "MDLATEX025", "MDLATEX026" })
            Assert.Contains(resultReport.Diagnostics, diagnostic => diagnostic.Code == code);
        Assert.Contains("Retained title", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains(@"\texttt{[x]}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains(@"\texttt{[ ]}", resultReport.GeneratedSource, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(@"\begin{theorem}\label{thm:x}Body\end{theorem} See \ref{thm:x}", "Body")]
    [InlineData(@"\begin{table}\label{tab:x}\begin{tabular}{l}Cell\\\end{tabular}\end{table} See \ref{tab:x}", "Cell")]
    [InlineData(@"\begin{figure}\label{fig:x}\caption{Figure marker}\end{figure} See \ref{fig:x}", "Figure marker")]
    [InlineData(@"\label{p:x}Body. See \ref{p:x}", "Body")]
    public void NonHeadingLabelsProduceWorkingPdfDestinations(string body, string marker) {
        var document = LatexDocument.Parse(@"\documentclass{article}\begin{document}" + body + @"\end{document}");
        byte[] bytes = document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.Contains(marker, OfficeIMO.Pdf.PdfDocument.Load(new System.IO.MemoryStream(bytes)).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void MissingReferencesRetainVisibleTextWithAnExportDiagnostic() {
        var document = LatexDocument.Parse(@"\documentclass{article}\begin{document}See \ref{missing:x}\end{document}");
        var result = document.ToPdf(null, out OfficeConversionReport resultReport);
        byte[] bytes = result.ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.Contains("missing:x", OfficeIMO.Pdf.PdfDocument.Load(new System.IO.MemoryStream(bytes)).ExtractText(), StringComparison.Ordinal);
        Assert.Contains(resultReport.Diagnostics, warning => warning.Code == "UnresolvedInternalLink");
    }

    [Theory]
    [InlineData(@"TRAILING\section{Inactive}")]
    [InlineData(@"TRAILING\[inactive\]")]
    public void ProjectionDoesNotExposeInertSourceAfterDocumentEnd(string trailer) {
        var document = LatexDocument.Parse(@"\documentclass{article}\begin{document}Public\end{document}" + trailer);
        Assert.Equal("Public", Assert.IsType<MarkdownTextRun>(Assert.Single(Assert.Single(document.ToMarkdown().Blocks.OfType<ParagraphBlock>()).Inlines.Nodes)).Text);
        string text = OfficeIMO.Pdf.PdfDocument.Load(new System.IO.MemoryStream(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).ExtractText();
        Assert.Contains("Public", text, StringComparison.Ordinal);
        Assert.DoesNotContain("TRAILING", text, StringComparison.Ordinal);
    }
}
