namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexConversionRegressionTests {
    [Fact]
    public void CaptionedTable_RemainsVisibleAndCarriesCaptionAndLabelMetadata() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\begin{table}\n" +
            "\\caption{Important values}\n" +
            "\\label{tab:values}\n" +
            "\\begin{tabular}{ll}\nA & B\\\\\nC & D\\\\\n\\end{tabular}\n" +
            "\\end{table}\n" +
            "\\end{document}\n";

        LatexMarkdownConversionResult result = LatexDocument.Parse(source).Document.ToMarkdownDocument();

        TableBlock table = Assert.Single(result.Document.Blocks.OfType<TableBlock>());
        Assert.Equal("tab:values", table.Attributes.ElementId);
        Assert.Equal("Important values", table.Attributes.GetAttribute("caption"));
        Assert.Contains("A", result.Document.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, static diagnostic =>
            diagnostic.Feature == "table-header" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Simplified);
    }

    [Fact]
    public void BracedListItem_ConvertsItsVisibleContent() {
        const string source = "\\documentclass{article}\n\\begin{document}\n\\begin{itemize}\\item {Visible item}\\end{itemize}\n\\end{document}\n";

        LatexMarkdownConversionResult result = LatexDocument.Parse(source).Document.ToMarkdownDocument();

        UnorderedListBlock list = Assert.Single(result.Document.Blocks.OfType<UnorderedListBlock>());
        Assert.Single(list.Items);
        Assert.Contains("- Visible item", result.Document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void FrontMatterTitle_DoesNotConsumeADifferentFirstHeading() {
        MarkdownDoc document = MarkdownReader.Parse("---\ntitle: Document title\n---\n\n# Introduction\n\nBody\n");

        MarkdownLatexConversionResult result = document.ToLatexDocument();

        Assert.Contains("\\title{Document title}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\maketitle", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\section{Introduction}", result.Source, StringComparison.Ordinal);
        Assert.Contains("Body", result.Source, StringComparison.Ordinal);
    }

    [Fact]
    public void TheoremCallouts_DeclareGeneratedTheoremEnvironments() {
        MarkdownDoc document = MarkdownDoc.Create().Callout("theorem", "Result", "Proof text.");

        MarkdownLatexConversionResult result = document.ToLatexDocument();

        Assert.Contains("\\usepackage{amsthm}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\newtheorem{theorem}{Theorem}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\begin{theorem}[Result]", result.Source, StringComparison.Ordinal);
    }

    [Fact]
    public void CanonicalArguments_EscapeTeXSpecialCharactersAndLabelsDeterministically() {
        MarkdownDoc document = MarkdownReader.Parse(
            "## Heading {#section%231}\n\n[query](https://example.test/a%20b?q=x#part&v=1)\n",
            new MarkdownReaderOptions { GenericAttributes = true });

        MarkdownLatexConversionResult result = document.ToLatexDocument();

        Assert.Contains("https://example.test/a\\%20b?q=x\\#part\\&v=1", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\label{section_0025_231}", result.Source, StringComparison.Ordinal);
        Assert.Contains(result.Diagnostics, static diagnostic => diagnostic.Feature == "label");
    }

    [Fact]
    public void CombinedTableSpans_UseValidNestingAndLogicalColumnCount() {
        TableBlock table = Assert.Single(MarkdownReader.Parse("| H |\n| --- |\n| wide |\n").Blocks.OfType<TableBlock>());
        TableCell cell = table.GetCell(0, 0)!;
        cell.ColumnSpan = 2;
        cell.RowSpan = 2;

        MarkdownLatexConversionResult result = MarkdownDoc.Create().Add(table).ToLatexDocument();

        Assert.Contains("\\begin{tabular}{ll}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\multicolumn{2}{l}{\\multirow{2}{*}{wide}}", result.Source, StringComparison.Ordinal);
    }

    [Fact]
    public void AdjacentHeadingLabel_BecomesMetadataWithoutADuplicateAnchorParagraph() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\section{Start}\n\\label{sec:start}\nBody.\n" +
            "\\end{document}\n";

        LatexMarkdownConversionResult result = LatexDocument.Parse(source).Document.ToMarkdownDocument();

        HeadingBlock heading = Assert.Single(result.Document.Blocks.OfType<HeadingBlock>());
        Assert.Equal("sec:start", heading.Attributes.ElementId);
        Assert.DoesNotContain(result.Document.Blocks.OfType<ParagraphBlock>(),
            static paragraph => paragraph.Inlines.Nodes.OfType<HtmlRawInline>().Any());
        Assert.Contains(result.Document.Blocks.OfType<ParagraphBlock>(),
            static paragraph => paragraph.Inlines.Nodes.OfType<TextRun>().Any(text => text.Text.Contains("Body.", StringComparison.Ordinal)));
    }

    [Fact]
    public void FigureAndTableContainerResiduals_RemainVisibleAndDiagnosed() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\begin{figure}\n\\centering\n\\includegraphics{plot.png}\n\\caption{Plot}\n\\end{figure}\n" +
            "\\begin{table}\n\\centering\n\\begin{tabular}{l}\nA\\\\\n\\end{tabular}\n\\end{table}\n" +
            "\\end{document}\n";

        LatexMarkdownConversionResult result = LatexDocument.Parse(source).Document.ToMarkdownDocument();

        Assert.Single(result.Document.Blocks.OfType<ImageBlock>());
        Assert.Single(result.Document.Blocks.OfType<TableBlock>());
        Assert.Equal(2, result.Diagnostics.Count(static diagnostic =>
            diagnostic.Code == "LATEXMD298" && diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback));
        Assert.Equal(2, result.Document.Blocks.OfType<CodeBlock>().Count(static block =>
            block.Content.Contains("\\centering", StringComparison.Ordinal)));
    }

    [Fact]
    public void CommonTextScriptsStrikeAndLineBreaks_ConvertSemantically() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "Text \\textsuperscript{two} \\textsubscript{sub} \\sout{gone}\\newline Next\\linebreak[4]Done.\n" +
            "\\end{document}\n";

        LatexMarkdownConversionResult result = LatexDocument.Parse(source).Document.ToMarkdownDocument();
        ParagraphBlock paragraph = Assert.Single(result.Document.Blocks.OfType<ParagraphBlock>());

        Assert.Single(paragraph.Inlines.Nodes.OfType<SuperscriptSequenceInline>());
        Assert.Single(paragraph.Inlines.Nodes.OfType<SubscriptSequenceInline>());
        Assert.Single(paragraph.Inlines.Nodes.OfType<StrikethroughSequenceInline>());
        Assert.Equal(2, paragraph.Inlines.Nodes.OfType<HardBreakInline>().Count());
        Assert.DoesNotContain(result.Diagnostics, static diagnostic =>
            diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
    }

    [Fact]
    public void NestedMarkdownScriptsAndStrike_GenerateBoundedProfileCommands() {
        MarkdownReaderOptions options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.Superscript = true;
        options.Subscript = true;
        MarkdownDoc document = MarkdownReader.Parse("^up **two**^ ~sub *italic*~ ~~gone **bold**~~", options);

        MarkdownLatexConversionResult result = document.ToLatexDocument();

        Assert.Contains("\\textsuperscript{up \\textbf{two}}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\textsubscript{sub \\emph{italic}}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\usepackage{ulem}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\sout{gone \\textbf{bold}}", result.Source, StringComparison.Ordinal);
        Assert.DoesNotContain("\\usepackage{amsmath}", result.Source, StringComparison.Ordinal);
    }
}
