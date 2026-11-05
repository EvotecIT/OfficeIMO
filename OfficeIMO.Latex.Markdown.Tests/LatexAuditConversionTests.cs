using OfficeIMO.Pdf;

namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexAuditConversionTests {
    private static string Wrap(string body) => "\\documentclass{article}\n\\begin{document}\n" + body + "\n\\end{document}";

    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    [InlineData("\r")]
    public void ScalarAndMathCommentRemovalConsumesOnlyTheCommentsLineEnding(string ending) {
        string suffix = "% PRIVATE" + ending + "b";
        string body = "\\href{a" + suffix + "}{link} \\url{a" + suffix + "} \\includegraphics{a" + suffix + ".png}" +
            "\n\n$$a" + suffix + "$$";
        var result = LatexDocument.Parse("\\author{a" + suffix + "}" + Wrap(body)).ToMarkdown(null);
        ParagraphBlock paragraph = Assert.Single(result.Blocks.OfType<ParagraphBlock>());
        LinkInline[] links = paragraph.Inlines.Nodes.OfType<LinkInline>().ToArray();
        Assert.Equal(2, links.Length);
        Assert.All(links, static link => Assert.Equal("ab", link.Url));
        Assert.Equal("ab.png", Assert.Single(paragraph.Inlines.Nodes.OfType<ImageInline>()).Src);
        Assert.Equal("ab", Assert.Single(result.Blocks.OfType<SemanticFencedBlock>()).Content);
        Assert.DoesNotContain("PRIVATE", result.ToMarkdown(), StringComparison.Ordinal);
        Assert.Equal("ab", result.FindFrontMatterEntry("author")?.Value);
        MarkdownDocument twoLines = LatexDocument.Parse("\\author{a% note" + ending + ending + "b}" + Wrap("Visible")).ToMarkdown();
        Assert.Equal("a" + ending + "b", twoLines.FindFrontMatterEntry("author")?.Value);
    }

    [Fact]
    public void InlineVerbatimPercentDoesNotBecomeACommentOrSplitItsSyntaxSpan() {
        var result = LatexDocument.Parse(Wrap("\\verb|100%|\n\nAfter")).ToMarkdown(null);
        ParagraphBlock[] paragraphs = result.Blocks.OfType<ParagraphBlock>().ToArray();
        Assert.Equal(2, paragraphs.Length);
        Assert.Equal("100%", Assert.Single(paragraphs[0].Inlines.Nodes.OfType<CodeSpanInline>()).Text);
        Assert.Equal("After", PlainText(paragraphs[1].Inlines));
    }

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
        string markdown = document.ToMarkdown().ToMarkdown();
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
        var result = document.ToMarkdown(new LatexMarkdownExportOptions { IncludePreambleAsFrontMatter = false }, out OfficeIMO.OfficeConversionReport conversionReport3);
        var resultReport = (LatexToMarkdownReport)conversionReport3;
        Assert.Contains("IMPORTANT CONTENT", result.ToMarkdown(), StringComparison.Ordinal);
        if (preserveOnly) {
            CodeBlock fallback = Assert.Single(result.Blocks.OfType<CodeBlock>());
            Assert.Equal(source, fallback.Content);
            LatexMarkdownConversionDiagnostic diagnostic = Assert.Single(resultReport.NativeDiagnostics, static item => item.Code == "LATEXMD297");
            Assert.Equal(0, diagnostic.LatexSpan!.Value.Start.Offset);
            Assert.Equal(source.Length, diagnostic.LatexSpan.Value.End.Offset);
        }
        Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
        Assert.ThrowsAny<OfficeConversionException>(() => resultReport.ThrowIfLoss());
        var omitted = document.ToMarkdown(new LatexMarkdownExportOptions { PreserveUnsupportedAsSource = false }, out OfficeIMO.OfficeConversionReport conversionReport4);
        var omittedReport = (LatexToMarkdownReport)conversionReport4;
        Assert.Contains(omittedReport.NativeDiagnostics, static diagnostic => diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
    }

    [Fact]
    public void UnsupportedFloatingTableContentIsVisibleAndReported() {
        var result = LatexDocument.Parse(Wrap("\\begin{table}\\begin{tabularx}{\\linewidth}{ll}IMPORTANT&B\\\\\\end{tabularx}\\end{table}")).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport5);
        var resultReport = (LatexToMarkdownReport)conversionReport5;
        Assert.Contains("IMPORTANT", result.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Feature == "environment:table");
    }

    [Fact]
    public void QuotesProjectFormattingAndSuppressCommentEnvironmentsAndLineComments() {
        var result = LatexDocument.Parse(Wrap("\\begin{quote}Public \\textbf{bold}% PRIVATE LINE\njoined" +
            "\\begin{comment}PRIVATE DRAFT\\end{comment}\\end{quote}")).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport6);
        var resultReport = (LatexToMarkdownReport)conversionReport6;
        string markdown = result.ToMarkdown();
        Assert.Contains("**bold**joined", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE", markdown, StringComparison.Ordinal);
        Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Feature == "comment-environment" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
    }

    [Fact]
    public void GeneratedLiteralCodeUnderlineAndLinkTextAndDestinationsRoundTrip() {
        const string literal = "a_b%{c}~^\\$&#";
        const string url = "https://example.test/a%20b?q=x#part&v=1";
        var inlines = new InlineSequence().Code(literal).Underline(literal).Link(literal, url);
        MarkdownDocument original = MarkdownDocument.Create();
        original.Add(new ParagraphBlock(inlines));
        var back = original.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport7).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport8);
        var backReport = (LatexToMarkdownReport)conversionReport8;
        InlineSequence converted = Assert.Single(back.Blocks.OfType<ParagraphBlock>()).Inlines;
        Assert.Equal(literal, Assert.Single(converted.Nodes.OfType<CodeSpanInline>()).Text);
        Assert.Equal(literal, Assert.Single(converted.Nodes.OfType<UnderlineInline>()).Text);
        LinkInline link = Assert.Single(converted.Nodes.OfType<LinkInline>());
        Assert.Equal(literal, link.Text);
        Assert.Equal(url, link.Url);
        Assert.False(backReport.HasLoss);
    }

    [Fact]
    public void BracketsInGeneratedOptionalTitlesAndTermsRemainInsideTheirArgument() {
        const string text = "Title] AFTER [more]";
        var definitions = new DefinitionListBlock();
        definitions.AddEntry(new DefinitionListEntry(new InlineSequence().Text(text), new IMarkdownBlock[] { new ParagraphBlock(new InlineSequence().Text("Definition body")) }));
        MarkdownDocument markdown = MarkdownDocument.Create();
        markdown.Callout("theorem", text, "Theorem body");
        markdown.Add(definitions);
        var latex = markdown.ToLatex(null);
        Assert.DoesNotContain(latex.Diagnostics, static diagnostic => diagnostic.Severity == LatexDiagnosticSeverity.Error);
        var result = latex.ToMarkdown(null);
        Assert.Equal(text, Assert.Single(result.Blocks.OfType<CalloutBlock>()).Title);
        DefinitionListEntry entry = Assert.Single(Assert.Single(result.Blocks.OfType<DefinitionListBlock>()).Entries);
        Assert.Equal(text, PlainText(entry.Term));
        Assert.Equal("Definition body", PlainText(Assert.IsType<ParagraphBlock>(Assert.Single(entry.DefinitionBlocks)).Inlines));
    }

    [Theory]
    [InlineData("é", "_00E9_")]
    [InlineData("é00E9_", "_00E9é")]
    [InlineData("_", "_005F_")]
    public void EncodedLabelNamesCannotCollideWithLiteralEscapeLookingIdentifiers(string first, string second) {
        MarkdownDocument markdown = MarkdownDocument.Parse("## A {#" + first + "}\n\n## B {#" + second + "}\n\n[A](#" + first + ") [B](#" + second + ")\n", new MarkdownLoadOptions { GenericAttributes = true });
        var result = markdown.ToLatex(new MarkdownLatexExportOptions { FirstHeadingIsTitle = false });
        string[] names = result.Labels.Select(static label => label.Name).ToArray();
        Assert.Equal(2, names.Length);
        Assert.NotEqual(names[0], names[1]);
        Assert.Equal(names.Select(static name => "#" + name), result.ToMarkdown().Descendants().OfType<LinkInline>().Select(static link => link.Url));
    }

    [Theory]
    [InlineData("\\pageref{sec:item}", "reference:pageref")]
    [InlineData("\\eqref{eq:item}", "reference:eqref")]
    [InlineData("\\includegraphics[width=2cm,angle=90]{plot.png}", "graphics-options")]
    [InlineData("\\textbf\n\n", "command-arguments:textbf")]
    [InlineData("\\textbf }", "command-arguments:textbf")]
    public void UnsupportedLayoutAndArgumentSemanticsAreReported(string body, string feature) {
        var result = LatexDocument.Parse(Wrap(body)).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport12);
        var resultReport = (LatexToMarkdownReport)conversionReport12;
        Assert.Contains(resultReport.NativeDiagnostics, diagnostic => diagnostic.Feature == feature);
        Assert.True(resultReport.HasLoss);
    }

    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    [InlineData("\r")]
    public void PercentCommentsConsumeTheirLineEndingAndPreserveFollowingBlankParagraphs(string ending) {
        LatexDocument document = LatexDocument.Parse(Wrap("word% note" + ending + "join\n\nnext% note" + ending + " \t" + ending + "line"));
        Assert.Equal(3, document.Paragraphs.Count);
        MarkdownDocument markdown = document.ToMarkdown();
        ParagraphBlock[] paragraphs = markdown.Blocks.OfType<ParagraphBlock>().ToArray();
        Assert.Equal(3, paragraphs.Length);
        Assert.Equal("wordjoin", PlainText(paragraphs[0].Inlines));
        Assert.Equal("next", PlainText(paragraphs[1].Inlines));
        Assert.Equal("line", PlainText(paragraphs[2].Inlines));
    }

    [Theory]
    [InlineData("\\texttt{\\textbf{important}}", "texttt")]
    [InlineData("\\underline{\\emph{important}}", "underline")]
    [InlineData("\\texttt{\\href{https://example.test}{important}}", "texttt")]
    public void ScalarInlineNodesReportFlattenedChildFormatting(string source, string command) {
        var result = LatexDocument.Parse(Wrap(source)).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport13);
        var resultReport = (LatexToMarkdownReport)conversionReport13;
        Assert.Equal("important", PlainText(Assert.Single(result.Blocks.OfType<ParagraphBlock>()).Inlines));
        Assert.Contains(resultReport.NativeDiagnostics, diagnostic => diagnostic.Code == "LATEXMD115" && diagnostic.Feature == "inline-formatting:" + command);
        Assert.ThrowsAny<OfficeConversionException>(() => resultReport.ThrowIfLoss());
    }

    [Theory]
    [InlineData("\\newcommand{\\showtitle}{\\maketitle}")]
    [InlineData("\\renewcommand{\\showtitle}{\\begin{verbatim}Inactive\\end{verbatim}}")]
    [InlineData("\\providecommand{\\showtitle}{\\verb|Inactive|}")]
    [InlineData("\\newtheorem{name}{\\maketitle}")]
    public void InactiveDefinitionChildrenDoNotSplitOrReplaceTheVisibleDefinitionFallback(string definition) {
        LatexDocument document = LatexDocument.Parse(Wrap(definition + "Actual"));
        Assert.Equal(definition + "Actual", Assert.Single(document.Paragraphs).Content);
        var result = document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport14);
        var resultReport = (LatexToMarkdownReport)conversionReport14;
        ParagraphBlock paragraph = Assert.Single(result.Blocks.OfType<ParagraphBlock>());
        Assert.Equal(definition, Assert.Single(paragraph.Inlines.Nodes.OfType<CodeSpanInline>()).Text);
        Assert.EndsWith("Actual", PlainText(paragraph.Inlines), StringComparison.Ordinal);
        Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
    }

    [Theory]
    [InlineData("\\newcommand{\\draft}{Public\\begin{comment}PRIVATE DRAFT\\end{comment}}")]
    [InlineData("\\custom{Public\\begin{comment}PRIVATE DRAFT\\end{comment}}")]
    [InlineData("\\href{Public\\begin{comment}PRIVATE DRAFT\\end{comment}}")]
    [InlineData("\\newcommand{\\draft}{Public% PRIVATE LINE\n}")]
    public void CommandFallbacksSuppressActualCommentSyntaxAndKeepTheEnclosingDefinition(string source) {
        LatexDocument document = LatexDocument.Parse(Wrap(source + "\n\nActual"));
        var result = document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport15);
        var resultReport = (LatexToMarkdownReport)conversionReport15;
        string markdown = result.ToMarkdown();
        Assert.DoesNotContain("PRIVATE", markdown, StringComparison.Ordinal);
        Assert.Contains("Public", markdown, StringComparison.Ordinal);
        Assert.Contains("Actual", markdown, StringComparison.Ordinal);
        Assert.Contains("PRIVATE", document.ToLatex(), StringComparison.Ordinal);
        Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
        if (source.IndexOf("\\begin{comment}", StringComparison.Ordinal) >= 0) {
            Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Feature == "comment-environment" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
        }
        Assert.DoesNotContain("PRIVATE", PdfDocument.Load(new System.IO.MemoryStream(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("\\title{Public\\begin{comment}PRIVATE\\end{comment}}", "Visible")]
    [InlineData("", "\\begin{figure}\\includegraphics{plot.png}\\caption{Public\\begin{comment}PRIVATE\\end{comment}}\\end{figure}")]
    [InlineData("", "\\begin{table}\\caption{Public\\begin{comment}PRIVATE\\end{comment}}\\begin{tabular}{l}Value\\\\\\end{tabular}\\end{table}")]
    [InlineData("", "\\begin{theorem}[Public\\begin{comment}PRIVATE\\end{comment}]Visible\\end{theorem}")]
    public void ScalarMetadataAndCaptionsSuppressActualCommentEnvironments(string preamble, string body) {
        var result = LatexDocument.Parse(preamble + Wrap(body)).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport16);
        var resultReport = (LatexToMarkdownReport)conversionReport16;
        Assert.Contains("Public", result.ToMarkdown(), StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE", result.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Feature == "comment-environment" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
    }

    [Theory]
    [InlineData("itemize")]
    [InlineData("enumerate")]
    public void CustomListLabelsRemainVisibleAndTheirMarkerSimplificationIsReported(string environment) {
        LatexDocument document = LatexDocument.Parse(Wrap("\\begin{" + environment + "}\\item[\\textbf{URGENT}] Call now\\end{" + environment + "}"));
        document.Lists[0].Items[0].Label = "\\textbf{EDITED}";
        var result = document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport17);
        var resultReport = (LatexToMarkdownReport)conversionReport17;
        Assert.Contains("**EDITED**: Call now", result.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Code == "LATEXMD214" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Simplified);
    }

    [Fact]
    public void PdfProjectionPreservesPlainSourceAndReportsItsFallback() {
        using var conversion = LatexDocument.Parse("IMPORTANT CONTENT").ToPdf(null, out OfficeConversionReport report);
        Assert.True(report.HasLoss);
        Assert.Contains(report.Diagnostics, static diagnostic => diagnostic.Code == "LATEXMD297");
        Assert.Contains("IMPORTANT CONTENT", PdfDocument.Load(new System.IO.MemoryStream(conversion.ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PreservationOnlyFallbackIncludesPreambleAndTrailerAndReportsWholeSourceOmission() {
        const string source = "\\documentclass{article}\n\\usepackage{custom}\n\\newcommand{\\greet}{Hello}\n" +
            "\\begin{document}\n\\greet World\n\\end{document}\nTRAILER";
        LatexDocument document = LatexDocument.Parse(source, new LatexParseOptions { Profile = LatexDocumentProfile.PreserveOnly });
        var preserved = document.ToMarkdown(new LatexMarkdownExportOptions { IncludePreambleAsFrontMatter = false });
        Assert.Equal(source, Assert.Single(preserved.Blocks.OfType<CodeBlock>()).Content);
        var pdf = document.ToPdf(null, out OfficeConversionReport pdfReport);
        string extracted = PdfDocument.Load(new System.IO.MemoryStream(pdf.ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).ExtractText();
        foreach (string text in new[] { "usepackage", "newcommand", "greet World", "TRAILER" }) Assert.Contains(text, extracted, StringComparison.Ordinal);
        var omitted = document.ToMarkdown(new LatexMarkdownExportOptions { IncludePreambleAsFrontMatter = false, PreserveUnsupportedAsSource = false }, out OfficeIMO.OfficeConversionReport conversionReport19);
        var omittedReport = (LatexToMarkdownReport)conversionReport19;
        Assert.Empty(omitted.Blocks);
        LatexMarkdownConversionDiagnostic diagnostic = Assert.Single(omittedReport.NativeDiagnostics, static item => item.Code == "LATEXMD297");
        Assert.Equal(LatexMarkdownConversionOutcome.Omitted, diagnostic.Outcome);
        Assert.Equal(0, diagnostic.LatexSpan!.Value.Start.Offset);
        Assert.Equal(source.Length, diagnostic.LatexSpan.Value.End.Offset);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PreservationOnlyFallbackKeepsOuterWhitespaceWithOrWithoutDocumentBody(bool body) {
        string source = " \n" + (body ? "\\begin{document}Body\\end{document}" : "\\newcommand{\\greet}{Hello}") + "\n ";
        MarkdownDocument result = LatexDocument.Parse(source, new LatexParseOptions { Profile = LatexDocumentProfile.PreserveOnly })
            .ToMarkdown(new LatexMarkdownExportOptions { IncludePreambleAsFrontMatter = false });
        Assert.Equal(source, Assert.Single(result.Blocks.OfType<CodeBlock>()).Content);
    }
    [Fact]
    public void PdfReportsParserErrorsIntroducedByNativeEdits() {
        LatexDocument document = LatexDocument.Parse(Wrap("Original paragraph"));
        document.Paragraphs[0].Content = "$unterminated";
        var conversion = document.ToPdf(null, out OfficeConversionReport conversionReport);
        Assert.Contains(conversionReport.Diagnostics, static warning => warning.Code == "LATEX003");
        Assert.True(conversionReport.HasLoss);
    }

    private static string PlainText(InlineSequence inlines) {
        var text = new StringBuilder();
        foreach (IPlainTextMarkdownInline inline in inlines.Nodes.OfType<IPlainTextMarkdownInline>()) inline.AppendPlainText(text);
        return text.ToString();
    }

}
