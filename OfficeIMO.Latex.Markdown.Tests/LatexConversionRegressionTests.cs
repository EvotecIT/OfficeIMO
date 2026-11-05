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

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport1);
        var resultReport = (LatexToMarkdownReport)conversionReport1;
        HeadingBlock[] converted = result.Blocks.OfType<HeadingBlock>().ToArray();

        Assert.Equal(new[] { firstLevel, secondLevel }, converted.Select(heading => heading.Level));
        Assert.Equal(2, resultReport.NativeDiagnostics.Count(diagnostic => diagnostic.Feature == "heading-numbering"));
    }

    [Theory]
    [InlineData("article", 2, 3)]
    [InlineData("report", 3, 4)]
    [InlineData("book", 3, 4)]
    public void PartReservesTheTopMarkdownHeadingLevel(string documentClass, int sectionLevel, int subsectionLevel) {
        string source = "\\documentclass{" + documentClass + "}\n\\begin{document}\n" +
            "\\part{Part}\\section{Section}\\subsection{Subsection}\n\\end{document}\n";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null);
        HeadingBlock[] headings = result.Blocks.OfType<HeadingBlock>().ToArray();

        Assert.Equal(new[] { "Part", "Section", "Subsection" }, headings.Select(heading => heading.Text));
        Assert.Equal(new[] { 1, sectionLevel, subsectionLevel }, headings.Select(heading => heading.Level));
    }

    [Fact]
    public void Starred_Heading_State_Is_Exposed_And_Reported_As_Simplified() {
        const string source = "\\documentclass{article}\n\\begin{document}\n\\section*{Unnumbered}\n\\end{document}\n";

        LatexDocument document = LatexDocument.ParseResult(source).Document;
        var result = document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport3);
        var resultReport = (LatexToMarkdownReport)conversionReport3;

        Assert.True(Assert.Single(document.Headings).IsStarred);
        Assert.Contains(resultReport.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD212" &&
            diagnostic.Message.IndexOf("unnumbered", StringComparison.Ordinal) >= 0);
        Assert.Equal(1, Assert.Single(result.Blocks.OfType<HeadingBlock>()).Level);
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

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport4);
        var resultReport = (LatexToMarkdownReport)conversionReport4;

        TableBlock table = Assert.Single(result.Blocks.OfType<TableBlock>());
        Assert.Equal("tab:values", table.Attributes.ElementId);
        Assert.Equal("Important values", table.Attributes.GetAttribute("caption"));
        Assert.Contains("A", ((result.ToMarkdown())), StringComparison.Ordinal);
        Assert.Contains(resultReport.NativeDiagnostics, static diagnostic =>
            diagnostic.Feature == "table-header" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Simplified);
    }

    [Fact]
    public void BracedListItem_ConvertsItsVisibleContent() {
        const string source = "\\documentclass{article}\n\\begin{document}\n\\begin{itemize}\\item {Visible item}\\end{itemize}\n\\end{document}\n";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null);

        UnorderedListBlock list = Assert.Single(result.Blocks.OfType<UnorderedListBlock>());
        Assert.Single(list.Items);
        Assert.Contains("- Visible item", ((result.ToMarkdown())), StringComparison.Ordinal);
    }

    [Fact]
    public void FrontMatterTitle_DoesNotConsumeADifferentFirstHeading() {
        MarkdownDocument document = MarkdownDocument.Parse("---\ntitle: Document title\n---\n\n# Introduction\n\nBody\n");

        var result = document.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport6);
        var resultReport = (MarkdownToLatexReport)conversionReport6;

        Assert.Contains("\\title{Document title}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\maketitle", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\section{Introduction}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("Body", resultReport.GeneratedSource, StringComparison.Ordinal);
    }

    [Fact]
    public void TheoremCallouts_DeclareGeneratedTheoremEnvironments() {
        MarkdownDocument document = MarkdownDocument.Create().Callout("theorem", "Result", "Proof text.");

        var result = document.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport7);
        var resultReport = (MarkdownToLatexReport)conversionReport7;

        Assert.Contains("\\usepackage{amsthm}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\newtheorem{theorem}{Theorem}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\begin{theorem}[Result]", resultReport.GeneratedSource, StringComparison.Ordinal);
    }

    [Fact]
    public void CanonicalArguments_EscapeTeXSpecialCharactersAndLabelsDeterministically() {
        MarkdownDocument document = MarkdownDocument.Parse(
            "## Heading {#section%231}\n\n[query](https://example.test/a%20b?q=x#part&v=1)\n",
            new MarkdownLoadOptions { GenericAttributes = true });

        var result = document.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport8);
        var resultReport = (MarkdownToLatexReport)conversionReport8;

        Assert.Contains("https://example.test/a\\%20b?q=x\\#part\\&v=1", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\label{section_0025_231}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Feature == "label");
    }

    [Fact]
    public void CombinedTableSpans_UseValidNestingAndLogicalColumnCount() {
        TableBlock table = Assert.Single(MarkdownDocument.Parse("| H |\n| --- |\n| wide |\n").Blocks.OfType<TableBlock>());
        TableCell cell = table.GetCell(0, 0)!;
        cell.ColumnSpan = 2;
        cell.RowSpan = 2;

        using var addedDocument = MarkdownDocument.Create();
        addedDocument.Add(table);
        var result = addedDocument.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport9);
        var resultReport = (MarkdownToLatexReport)conversionReport9;

        Assert.Contains("\\begin{tabular}{ll}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\multicolumn{2}{l}{\\multirow{2}{*}{wide}}", resultReport.GeneratedSource, StringComparison.Ordinal);
    }

    [Fact]
    public void AdjacentHeadingLabel_BecomesMetadataWithoutADuplicateAnchorParagraph() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\section{Start}\n\\label{sec:start}\nBody.\n" +
            "\\end{document}\n";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null);

        HeadingBlock heading = Assert.Single(result.Blocks.OfType<HeadingBlock>());
        Assert.Equal("sec:start", heading.Attributes.ElementId);
        Assert.DoesNotContain(result.Blocks.OfType<ParagraphBlock>(),
            static paragraph => paragraph.Inlines.Nodes.OfType<HtmlRawInline>().Any());
        Assert.Contains(result.Blocks.OfType<ParagraphBlock>(),
            static paragraph => paragraph.Inlines.Nodes.OfType<MarkdownTextRun>().Any(text => text.Text.IndexOf("Body.", StringComparison.Ordinal) >= 0));
    }

    [Fact]
    public void FigureAndTableContainerResiduals_RemainVisibleAndDiagnosed() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\begin{figure}\n\\centering\n\\includegraphics{plot.png}\n\\caption{Plot}\n\\end{figure}\n" +
            "\\begin{table}\n\\centering\n\\begin{tabular}{l}\nA\\\\\n\\end{tabular}\n\\end{table}\n" +
            "\\end{document}\n";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport11);
        var resultReport = (LatexToMarkdownReport)conversionReport11;

        Assert.Single(result.Blocks.OfType<ImageBlock>());
        Assert.Single(result.Blocks.OfType<TableBlock>());
        Assert.Equal(2, resultReport.NativeDiagnostics.Count(static diagnostic =>
            diagnostic.Code == "LATEXMD298" && diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback));
        Assert.Equal(2, result.Blocks.OfType<CodeBlock>().Count(static block =>
            block.Content.IndexOf("\\centering", StringComparison.Ordinal) >= 0));
    }

    [Fact]
    public void CommonTextScriptsStrikeAndLineBreaks_ConvertSemantically() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "Text \\textsuperscript{two} \\textsubscript{sub} \\sout{gone}\\newline Next\\linebreak[4]Done.\n" +
            "\\end{document}\n";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport12);
        var resultReport = (LatexToMarkdownReport)conversionReport12;
        ParagraphBlock paragraph = Assert.Single(result.Blocks.OfType<ParagraphBlock>());

        Assert.Single(paragraph.Inlines.Nodes.OfType<SuperscriptSequenceInline>());
        Assert.Single(paragraph.Inlines.Nodes.OfType<SubscriptSequenceInline>());
        Assert.Single(paragraph.Inlines.Nodes.OfType<StrikethroughSequenceInline>());
        Assert.Equal(2, paragraph.Inlines.Nodes.OfType<HardBreakInline>().Count());
        Assert.DoesNotContain(resultReport.NativeDiagnostics, static diagnostic =>
            diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
    }

    [Fact]
    public void NestedMarkdownScriptsAndStrike_GenerateBoundedProfileCommands() {
        MarkdownLoadOptions options = MarkdownLoadOptions.CreateOfficeIMOProfile();
        options.Superscript = true;
        options.Subscript = true;
        MarkdownDocument document = MarkdownDocument.Parse("^up **two**^ ~sub *italic*~ ~~gone **bold**~~", options);

        var result = document.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport13);
        var resultReport = (MarkdownToLatexReport)conversionReport13;

        Assert.Contains("\\textsuperscript{up \\textbf{two}}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\textsubscript{sub \\emph{italic}}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\usepackage[normalem]{ulem}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\sout{gone \\textbf{bold}}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.DoesNotContain("\\usepackage{amsmath}", resultReport.GeneratedSource, StringComparison.Ordinal);
    }

    [Fact]
    public void CommentAndVerbatimEnvironmentsRemainOpaqueDuringMarkdownConversion() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "Before.\n\\begin{comment}SECRET \\section{Hidden}\\end{comment}\n" +
            "\\begin{verbatim}literal % value { \\command\\end{verbatim}\n" +
            "After.\n\\end{document}\n";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport14);
        var resultReport = (LatexToMarkdownReport)conversionReport14;
        string markdown = ((result.ToMarkdown()));

        Assert.DoesNotContain("SECRET", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden", markdown, StringComparison.Ordinal);
        Assert.Contains("literal % value { \\command", markdown, StringComparison.Ordinal);
        Assert.Contains(resultReport.NativeDiagnostics, diagnostic =>
            diagnostic.Code == "LATEXMD210" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
        Assert.Contains(resultReport.NativeDiagnostics, diagnostic =>
            diagnostic.Code == "LATEXMD213" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Simplified);
    }

    [Fact]
    public void VerbatimEnvironmentArgumentsAreNotEmittedAsCode() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\begin{minted}[linenos]{csharp}\nConsole.WriteLine(1);\n\\end{minted}\n" +
            "\\begin{lstlisting}[language=C]\nprintf(\"ok\");\n\\end{lstlisting}\n" +
            "\\begin {Verbatim}[numbers=left]\nopaque();\n\\end {Verbatim}\n" +
            "\\end{document}\n";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null);
        CodeBlock[] blocks = result.Blocks.OfType<CodeBlock>().ToArray();

        Assert.Equal(3, blocks.Length);
        Assert.Contains("Console.WriteLine(1);", blocks[0].Content, StringComparison.Ordinal);
        Assert.DoesNotContain("linenos", blocks[0].Content, StringComparison.Ordinal);
        Assert.DoesNotContain("{csharp}", blocks[0].Content, StringComparison.Ordinal);
        Assert.Contains("printf(\"ok\");", blocks[1].Content, StringComparison.Ordinal);
        Assert.DoesNotContain("language=C", blocks[1].Content, StringComparison.Ordinal);
        Assert.Contains("opaque();", blocks[2].Content, StringComparison.Ordinal);
        Assert.DoesNotContain("numbers=left", blocks[2].Content, StringComparison.Ordinal);
    }

    [Fact]
    public void MintedArgumentsAfterLineEndingsAreNotEmittedAsCode() {
        const string source =
            "\\documentclass{article}\r\n\\begin{document}\r\n" +
            "\\begin{minted}\r\n[linenos]\r\n{csharp}\r\nConsole.WriteLine(1);\r\n\\end{minted}\r\n" +
            "\\end{document}\r\n";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null);
        CodeBlock block = Assert.Single(result.Blocks.OfType<CodeBlock>());

        Assert.Contains("Console.WriteLine(1);", block.Content, StringComparison.Ordinal);
        Assert.DoesNotContain("linenos", block.Content, StringComparison.Ordinal);
        Assert.DoesNotContain("{csharp}", block.Content, StringComparison.Ordinal);
    }

    [Fact]
    public void InlineVerbPreservesCommentMarkersAsCode() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\nBefore \\verb|a%b{c}| after.\n\\end{document}\n";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport17);
        var resultReport = (LatexToMarkdownReport)conversionReport17;
        string markdown = ((result.ToMarkdown()));

        Assert.Contains("`a%b{c}`", markdown, StringComparison.Ordinal);
        Assert.Contains("after", markdown, StringComparison.Ordinal);
        Assert.Contains(resultReport.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD111");
    }

    [Fact]
    public void CommentNestedInsideUnknownEnvironmentIsOmittedFromSourceFallback() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\n" +
            "\\begin{unknown}Visible before.\\begin{comment}SECRET\\section{Hidden}\\end{comment}Visible after.\\end{unknown}\n" +
            "\\end{document}\n";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport18);
        var resultReport = (LatexToMarkdownReport)conversionReport18;
        string markdown = ((result.ToMarkdown()));

        Assert.Contains("Visible before", markdown, StringComparison.Ordinal);
        Assert.Contains("Visible after", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("SECRET", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden", markdown, StringComparison.Ordinal);
        Assert.Contains(resultReport.NativeDiagnostics, diagnostic => diagnostic.Code == "LATEXMD210" &&
            diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
        Assert.Contains(resultReport.NativeDiagnostics, diagnostic => diagnostic.Code == "LATEXMD299" &&
            diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
    }

    [Fact]
    public void UnterminatedInlineVerbPreservesEveryContentCharacter() {
        const string source = "\\documentclass{article}\n\\begin{document}\nBefore \\verb|abc";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport19);
        var resultReport = (LatexToMarkdownReport)conversionReport19;

        Assert.Contains("`abc`", ((result.ToMarkdown())), StringComparison.Ordinal);
        Assert.Contains(resultReport.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD111");
    }

    [Fact]
    public void UnterminatedInlineVerbDoesNotConsumeFollowingLines() {
        const string source =
            "\\documentclass{article}\n\\begin{document}\nBefore \\verb|abc\n\\section{Next}\nAfter\n\\end{document}\n";

        var result = LatexDocument.ParseResult(source).Document.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport20);
        var resultReport = (LatexToMarkdownReport)conversionReport20;
        string markdown = ((result.ToMarkdown()));

        Assert.Contains("`abc`", markdown, StringComparison.Ordinal);
        Assert.Contains("# Next", markdown, StringComparison.Ordinal);
        Assert.Contains("After", markdown, StringComparison.Ordinal);
        Assert.Contains(resultReport.Diagnostics, diagnostic => diagnostic.Code == "LATEXMD111");
    }
}
