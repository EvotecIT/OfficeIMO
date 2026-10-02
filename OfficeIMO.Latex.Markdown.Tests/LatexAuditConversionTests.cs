using OfficeIMO.Pdf;

namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexAuditConversionTests {
    private static string Wrap(string body) => "\\documentclass{article}\n\\begin{document}\n" + body + "\n\\end{document}";

    [Fact]
    public void ConversionUsesEditedHeadingParagraphCellListAndMathContent() {
        LatexDocument document = LatexDocument.Parse(Wrap("\\section{Old heading}\nOld paragraph\n\n" +
            "\\begin{tabular}{ll}Old cell&B\\\\\\end{tabular}\n" +
            "\\begin{itemize}\\item Old item\\end{itemize}\n$$old math$$"));
        document.Headings[0].Title = "New heading";
        document.Paragraphs[0].Content = "New paragraph";
        document.Tables[0].Rows[0].Cells[0].Content = "New cell";
        document.Lists[0].Items[0].Content = "New item";
        document.Math[0].Content = "new math";
        string markdown = document.ToMarkdownDocument().ToMarkdown();
        foreach (string content in new[] { "New heading", "New paragraph", "New cell", "New item", "new math" }) Assert.Contains(content, markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("Old", markdown, StringComparison.Ordinal);
        Assert.Contains("Old heading", document.Source.Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PlainAndPreservationOnlySourcesProduceDiagnosedFallback(bool preserveOnly) {
        string source = preserveOnly ? Wrap("IMPORTANT CONTENT") : "IMPORTANT CONTENT";
        LatexDocument document = LatexDocument.Parse(source, new LatexParseOptions {
            Profile = preserveOnly ? LatexDocumentProfile.PreserveOnly : LatexDocumentProfile.OfficeIMO
        });
        LatexToMarkdownResult result = document.ToMarkdownDocumentResult(new LatexToMarkdownOptions { IncludePreambleAsFrontMatter = false });
        Assert.Contains("IMPORTANT CONTENT", result.Value.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, static diagnostic => diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        LatexToMarkdownResult omitted = document.ToMarkdownDocumentResult(new LatexToMarkdownOptions { PreserveUnsupportedAsSource = false });
        Assert.Contains(omitted.Report.Diagnostics, static diagnostic => diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
    }

    [Fact]
    public void UnsupportedFloatingTableContentIsVisibleAndReported() {
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap("\\begin{table}\\begin{tabularx}{\\linewidth}{ll}IMPORTANT&B\\\\\\end{tabularx}\\end{table}")).ToMarkdownDocumentResult();
        Assert.Contains("IMPORTANT", result.Value.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, static diagnostic => diagnostic.Feature == "environment:table");
    }

    [Fact]
    public void QuotesProjectFormattingAndSuppressCommentEnvironmentsAndLineComments() {
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap("\\begin{quote}Public \\textbf{bold}% PRIVATE LINE\njoined" +
            "\\begin{comment}PRIVATE DRAFT\\end{comment}\\end{quote}")).ToMarkdownDocumentResult();
        string markdown = result.Value.ToMarkdown();
        Assert.Contains("**bold**joined", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE", markdown, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, static diagnostic => diagnostic.Feature == "comment-environment" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
    }

    [Fact]
    public void GeneratedLiteralCodeUnderlineAndLinkTextAndDestinationsRoundTrip() {
        const string literal = "a_b%{c}~^\\$&#";
        const string url = "https://example.test/a%20b?q=x#part&v=1";
        var inlines = new InlineSequence().Code(literal).Underline(literal).Link(literal, url);
        MarkdownDoc original = MarkdownDoc.Create().Add(new ParagraphBlock(inlines));
        LatexToMarkdownResult back = original.ToLatexDocumentResult().Value.ToMarkdownDocumentResult();
        InlineSequence converted = Assert.Single(back.Value.Blocks.OfType<ParagraphBlock>()).Inlines;
        Assert.Equal(literal, Assert.Single(converted.Nodes.OfType<CodeSpanInline>()).Text);
        Assert.Equal(literal, Assert.Single(converted.Nodes.OfType<UnderlineInline>()).Text);
        LinkInline link = Assert.Single(converted.Nodes.OfType<LinkInline>());
        Assert.Equal(literal, link.Text);
        Assert.Equal(url, link.Url);
        Assert.False(back.Report.HasLoss);
    }

    [Fact]
    public void BracketsInGeneratedOptionalTitlesAndTermsRemainInsideTheirArgument() {
        const string text = "Title] AFTER [more]";
        var definitions = new DefinitionListBlock();
        definitions.AddEntry(new DefinitionListEntry(new InlineSequence().Text(text), new IMarkdownBlock[] { new ParagraphBlock(new InlineSequence().Text("Definition body")) }));
        MarkdownDoc markdown = MarkdownDoc.Create().Callout("theorem", text, "Theorem body").Add(definitions);
        MarkdownToLatexResult latex = markdown.ToLatexDocumentResult();
        Assert.DoesNotContain(latex.Value.Diagnostics, static diagnostic => diagnostic.Severity == LatexDiagnosticSeverity.Error);
        LatexToMarkdownResult result = latex.Value.ToMarkdownDocumentResult();
        Assert.Equal(text, Assert.Single(result.Value.Blocks.OfType<CalloutBlock>()).Title);
        DefinitionListEntry entry = Assert.Single(Assert.Single(result.Value.Blocks.OfType<DefinitionListBlock>()).Entries);
        Assert.Equal(text, PlainText(entry.Term));
        Assert.Equal("Definition body", PlainText(Assert.IsType<ParagraphBlock>(Assert.Single(entry.DefinitionBlocks)).Inlines));
    }

    [Fact]
    public void EncodedLabelNamesCannotCollideWithLiteralEscapeLookingIdentifiers() {
        MarkdownDoc markdown = MarkdownReader.Parse("## A {#é}\n\n## B {#_00E9_}\n\n[A](#é) [B](#_00E9_)\n", new MarkdownReaderOptions { GenericAttributes = true });
        MarkdownToLatexResult result = markdown.ToLatexDocumentResult(new MarkdownToLatexOptions { FirstHeadingIsTitle = false });
        string[] names = result.Value.Labels.Select(static label => label.Name).ToArray();
        Assert.Equal(2, names.Length);
        Assert.NotEqual(names[0], names[1]);
        Assert.Equal(names, result.Value.References.Select(static reference => reference.Target));
    }

    [Theory]
    [InlineData("\\pageref{sec:item}", "reference:pageref")]
    [InlineData("\\eqref{eq:item}", "reference:eqref")]
    [InlineData("\\includegraphics[width=2cm,angle=90]{plot.png}", "graphics-options")]
    [InlineData("\\textbf", "command-arguments:textbf")]
    [InlineData("\\textbf X", "command-arguments:textbf")]
    public void UnsupportedLayoutAndArgumentSemanticsAreReported(string body, string feature) {
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap(body)).ToMarkdownDocumentResult();
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Feature == feature);
        Assert.True(result.Report.HasLoss);
    }

    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    [InlineData("\r")]
    public void PercentCommentsConsumeTheirLineEndingWithoutCreatingSpaceOrParagraphs(string ending) {
        MarkdownDoc markdown = LatexDocument.Parse(Wrap("word% note" + ending + "join\n\nnext% note" + ending + ending + "line")).ToMarkdownDocument();
        ParagraphBlock[] paragraphs = markdown.Blocks.OfType<ParagraphBlock>().ToArray();
        Assert.Equal(2, paragraphs.Length);
        Assert.Equal("wordjoin", PlainText(paragraphs[0].Inlines));
        Assert.Equal("next\nline", PlainText(paragraphs[1].Inlines));
    }

    [Fact]
    public void PdfProjectionPreservesPlainSourceAndReportsItsFallback() {
        var conversion = LatexDocument.Parse("IMPORTANT CONTENT").ToPdfDocumentResult();
        Assert.True(conversion.HasLoss);
        Assert.Contains(conversion.FidelityDiagnostics, static diagnostic => diagnostic.Code == "LATEXMD297");
        Assert.Contains("IMPORTANT CONTENT", PdfReadDocument.Open(conversion.Value.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }
    [Fact]
    public void PdfReportsParserErrorsIntroducedByNativeEdits() {
        LatexDocument document = LatexDocument.Parse(Wrap("Original paragraph"));
        document.Paragraphs[0].Content = "$unterminated";
        var conversion = document.ToPdfDocumentResult();
        Assert.Contains(conversion.Warnings, static warning => warning.Code == "LATEX003");
        Assert.True(conversion.HasLoss);
    }

    private static string PlainText(InlineSequence inlines) {
        var text = new StringBuilder();
        foreach (IPlainTextMarkdownInline inline in inlines.Nodes.OfType<IPlainTextMarkdownInline>()) inline.AppendPlainText(text);
        return text.ToString();
    }

}
