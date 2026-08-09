namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexConversionRegressionTests {
    [Theory]
    [InlineData("article", "\\section{One}\n\\subsection{Two}", 1, 2)]
    [InlineData("report", "\\chapter{One}\n\\section{Two}", 1, 2)]
    [InlineData("book", "\\chapter{One}\n\\section{Two}", 1, 2)]
    public void Heading_Hierarchy_Is_Normalized_For_The_Document_Class(
        string documentClass,
        string headings,
        int firstLevel,
        int secondLevel) {
        string source = "\\documentclass{" + documentClass + "}\n\\begin{document}\n" + headings + "\n\\end{document}\n";

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();
        HeadingBlock[] converted = result.Value.Blocks.OfType<HeadingBlock>().ToArray();

        Assert.Equal(new[] { firstLevel, secondLevel }, converted.Select(heading => heading.Level));
        Assert.Equal(2, result.Report.Diagnostics.Count(diagnostic => diagnostic.Feature == "heading-numbering"));
    }

    [Fact]
    public void Starred_Heading_State_Is_Exposed_And_Reported_As_Simplified() {
        const string source = "\\documentclass{article}\n\\begin{document}\n\\section*{Unnumbered}\n\\end{document}\n";

        LatexDocument document = LatexDocument.Parse(source).Document;
        LatexToMarkdownResult result = document.ToMarkdownDocumentResult();

        Assert.True(Assert.Single(document.Headings).IsStarred);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD212" &&
            diagnostic.Message.Contains("unnumbered", StringComparison.Ordinal));
        Assert.Equal(1, Assert.Single(result.Value.Blocks.OfType<HeadingBlock>()).Level);
    }

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

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();

        TableBlock table = Assert.Single(result.Value.Blocks.OfType<TableBlock>());
        Assert.Equal("tab:values", table.Attributes.ElementId);
        Assert.Equal("Important values", table.Attributes.GetAttribute("caption"));
        Assert.Contains("A", result.Value.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, static diagnostic =>
            diagnostic.Feature == "table-header" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Simplified);
    }

    [Fact]
    public void BracedListItem_ConvertsItsVisibleContent() {
        const string source = "\\documentclass{article}\n\\begin{document}\n\\begin{itemize}\\item {Visible item}\\end{itemize}\n\\end{document}\n";

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();

        UnorderedListBlock list = Assert.Single(result.Value.Blocks.OfType<UnorderedListBlock>());
        Assert.Single(list.Items);
        Assert.Contains("- Visible item", result.Value.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void FrontMatterTitle_DoesNotConsumeADifferentFirstHeading() {
        MarkdownDoc document = MarkdownReader.Parse("---\ntitle: Document title\n---\n\n# Introduction\n\nBody\n");

        MarkdownToLatexResult result = document.ToLatexDocumentResult();

        Assert.Contains("\\title{Document title}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\maketitle", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\section{Introduction}", result.Source, StringComparison.Ordinal);
        Assert.Contains("Body", result.Source, StringComparison.Ordinal);
    }

    [Fact]
    public void TheoremCallouts_DeclareGeneratedTheoremEnvironments() {
        MarkdownDoc document = MarkdownDoc.Create().Callout("theorem", "Result", "Proof text.");

        MarkdownToLatexResult result = document.ToLatexDocumentResult();

        Assert.Contains("\\usepackage{amsthm}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\newtheorem{theorem}{Theorem}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\begin{theorem}[Result]", result.Source, StringComparison.Ordinal);
    }

    [Fact]
    public void CanonicalArguments_EscapeTeXSpecialCharactersAndLabelsDeterministically() {
        MarkdownDoc document = MarkdownReader.Parse(
            "## Heading {#section%231}\n\n[query](https://example.test/a%20b?q=x#part&v=1)\n",
            new MarkdownReaderOptions { GenericAttributes = true });

        MarkdownToLatexResult result = document.ToLatexDocumentResult();

        Assert.Contains("https://example.test/a\\%20b?q=x\\#part\\&v=1", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\label{section_0025_231}", result.Source, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, static diagnostic => diagnostic.Feature == "label");
    }

    [Fact]
    public void CombinedTableSpans_UseValidNestingAndLogicalColumnCount() {
        TableBlock table = Assert.Single(MarkdownReader.Parse("| H |\n| --- |\n| wide |\n").Blocks.OfType<TableBlock>());
        TableCell cell = table.GetCell(0, 0)!;
        cell.ColumnSpan = 2;
        cell.RowSpan = 2;

        MarkdownToLatexResult result = MarkdownDoc.Create().Add(table).ToLatexDocumentResult();

        Assert.Contains("\\begin{tabular}{ll}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\multicolumn{2}{l}{\\multirow{2}{*}{wide}}", result.Source, StringComparison.Ordinal);
    }

    [Fact]
    public void AdjacentHeadingLabel_BecomesMetadataWithoutADuplicateAnchorParagraph() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\section{Start}\n\\label{sec:start}\nBody.\n" +
            "\\end{document}\n";

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();

        HeadingBlock heading = Assert.Single(result.Value.Blocks.OfType<HeadingBlock>());
        Assert.Equal("sec:start", heading.Attributes.ElementId);
        Assert.DoesNotContain(result.Value.Blocks.OfType<ParagraphBlock>(),
            static paragraph => paragraph.Inlines.Nodes.OfType<HtmlRawInline>().Any());
        Assert.Contains(result.Value.Blocks.OfType<ParagraphBlock>(),
            static paragraph => paragraph.Inlines.Nodes.OfType<MarkdownTextRun>().Any(text => text.Text.Contains("Body.", StringComparison.Ordinal)));
    }

    [Fact]
    public void FigureAndTableContainerResiduals_RemainVisibleAndDiagnosed() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\begin{figure}\n\\centering\n\\includegraphics{plot.png}\n\\caption{Plot}\n\\end{figure}\n" +
            "\\begin{table}\n\\centering\n\\begin{tabular}{l}\nA\\\\\n\\end{tabular}\n\\end{table}\n" +
            "\\end{document}\n";

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();

        Assert.Single(result.Value.Blocks.OfType<ImageBlock>());
        Assert.Single(result.Value.Blocks.OfType<TableBlock>());
        Assert.Equal(2, result.Report.Diagnostics.Count(static diagnostic =>
            diagnostic.Code == "LATEXMD298" && diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback));
        Assert.Equal(2, result.Value.Blocks.OfType<CodeBlock>().Count(static block =>
            block.Content.Contains("\\centering", StringComparison.Ordinal)));
    }

    [Fact]
    public void CommonTextScriptsStrikeAndLineBreaks_ConvertSemantically() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "Text \\textsuperscript{two} \\textsubscript{sub} \\sout{gone}\\newline Next\\linebreak[4]Done.\n" +
            "\\end{document}\n";

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();
        ParagraphBlock paragraph = Assert.Single(result.Value.Blocks.OfType<ParagraphBlock>());

        Assert.Single(paragraph.Inlines.Nodes.OfType<SuperscriptSequenceInline>());
        Assert.Single(paragraph.Inlines.Nodes.OfType<SubscriptSequenceInline>());
        Assert.Single(paragraph.Inlines.Nodes.OfType<StrikethroughSequenceInline>());
        Assert.Equal(2, paragraph.Inlines.Nodes.OfType<HardBreakInline>().Count());
        Assert.DoesNotContain(result.Report.Diagnostics, static diagnostic =>
            diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
    }

    [Fact]
    public void NestedMarkdownScriptsAndStrike_GenerateBoundedProfileCommands() {
        MarkdownReaderOptions options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.Superscript = true;
        options.Subscript = true;
        MarkdownDoc document = MarkdownReader.Parse("^up **two**^ ~sub *italic*~ ~~gone **bold**~~", options);

        MarkdownToLatexResult result = document.ToLatexDocumentResult();

        Assert.Contains("\\textsuperscript{up \\textbf{two}}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\textsubscript{sub \\emph{italic}}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\usepackage{ulem}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\sout{gone \\textbf{bold}}", result.Source, StringComparison.Ordinal);
        Assert.DoesNotContain("\\usepackage{amsmath}", result.Source, StringComparison.Ordinal);
    }

    [Fact]
    public void CommentAndVerbatimEnvironmentsRemainOpaqueDuringMarkdownConversion() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "Before.\n\\begin{comment}SECRET \\section{Hidden}\\end{comment}\n" +
            "\\begin{verbatim}literal % value { \\command\\end{verbatim}\n" +
            "After.\n\\end{document}\n";

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();
        string markdown = result.Value.ToMarkdown();

        Assert.DoesNotContain("SECRET", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden", markdown, StringComparison.Ordinal);
        Assert.Contains("literal % value { \\command", markdown, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "LATEXMD210" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "LATEXMD213" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Simplified);
    }

    [Fact]
    public void VerbatimEnvironmentArgumentsAreNotEmittedAsCode() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\begin{minted}[linenos]{csharp}\nConsole.WriteLine(1);\n\\end{minted}\n" +
            "\\begin{lstlisting}[language=C]\nprintf(\"ok\");\n\\end{lstlisting}\n" +
            "\\end{document}\n";

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();
        CodeBlock[] blocks = result.Value.Blocks.OfType<CodeBlock>().ToArray();

        Assert.Equal(2, blocks.Length);
        Assert.Contains("Console.WriteLine(1);", blocks[0].Content, StringComparison.Ordinal);
        Assert.DoesNotContain("linenos", blocks[0].Content, StringComparison.Ordinal);
        Assert.DoesNotContain("{csharp}", blocks[0].Content, StringComparison.Ordinal);
        Assert.Contains("printf(\"ok\");", blocks[1].Content, StringComparison.Ordinal);
        Assert.DoesNotContain("language=C", blocks[1].Content, StringComparison.Ordinal);
    }

    [Fact]
    public void InlineVerbPreservesCommentMarkersAsCode() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\nBefore \\verb|a%b{c}| after.\n\\end{document}\n";

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();
        string markdown = result.Value.ToMarkdown();

        Assert.Contains("`a%b{c}`", markdown, StringComparison.Ordinal);
        Assert.Contains("after", markdown, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD111");
    }

    [Fact]
    public void CommentNestedInsideUnknownEnvironmentIsOmittedFromSourceFallback() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\begin{unknown}Visible before.\\begin{comment}SECRET\\section{Hidden}\\end{comment}Visible after.\\end{unknown}\n" +
            "\\end{document}\n";

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();
        string markdown = result.Value.ToMarkdown();

        Assert.Contains("Visible before", markdown, StringComparison.Ordinal);
        Assert.Contains("Visible after", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("SECRET", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden", markdown, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD210" &&
            diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD299" &&
            diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
    }

    [Fact]
    public void UnterminatedInlineVerbPreservesEveryContentCharacter() {
        const string source = "\\documentclass{article}\n\\begin{document}\nBefore \\verb|abc";

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();

        Assert.Contains("`abc`", result.Value.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD111");
    }

    [Fact]
    public void UnterminatedInlineVerbDoesNotConsumeFollowingLines() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\nBefore \\verb|abc\n\\section{Next}\nAfter\n\\end{document}\n";

        LatexToMarkdownResult result = LatexDocument.Parse(source).Document.ToMarkdownDocumentResult();
        string markdown = result.Value.ToMarkdown();

        Assert.Contains("`abc`", markdown, StringComparison.Ordinal);
        Assert.Contains("# Next", markdown, StringComparison.Ordinal);
        Assert.Contains("After", markdown, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD111");
    }
}
