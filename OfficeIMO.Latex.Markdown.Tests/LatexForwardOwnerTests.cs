using OfficeIMO.Pdf;

namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexForwardOwnerTests {
    private static string Wrap(string body) => "\\documentclass{article}\\begin{document}" + body + "\\end{document}";

    [Fact]
    public void PreservedPartsCannotShiftVisibleHeadingLevels() {
        MarkdownDoc result = LatexDocument.Parse(Wrap("\\unknown{\\part{Hidden}}\\section{Visible}")).ToMarkdownDocument();
        Assert.Equal(1, Assert.Single(result.Blocks.OfType<HeadingBlock>()).Level);
        Assert.Contains("\\unknown{\\part{Hidden}}", Assert.Single(result.Blocks.OfType<CodeBlock>()).Content, StringComparison.Ordinal);
    }

    [Fact]
    public void InertMetadataCannotShadowActivePreambleOrCreateProjectedMetadata() {
        const string trailer = "\\title{INERT TITLE}\\author{PRIVATE}\\date{INERT DATE}\\documentclass{report}";
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap("Public") + trailer).ToMarkdownDocumentResult();
        Assert.Null(result.Value.FindFrontMatterEntry("title"));
        Assert.Null(result.Value.FindFrontMatterEntry("author"));
        Assert.Null(result.Value.FindFrontMatterEntry("date"));
        Assert.Equal("article", result.Value.FindFrontMatterEntry("documentclass")?.Value);
        Assert.Empty(result.Value.Blocks.OfType<HeadingBlock>());
        string text = PdfReadDocument.Open(LatexDocument.Parse(Wrap("Public") + trailer).ToPdfBytes()).ExtractText();
        Assert.Contains("Public", text, StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE", text, StringComparison.Ordinal);
        Assert.DoesNotContain("INERT", text, StringComparison.Ordinal);
        MarkdownDoc active = LatexDocument.Parse("\\title{Active}\\author{Author}\\date{Today}" + Wrap("Body") + trailer).ToMarkdownDocument();
        Assert.Equal("Active", active.FindFrontMatterEntry("title")?.Value);
        Assert.Equal("Author", active.FindFrontMatterEntry("author")?.Value);
        Assert.Equal("Today", active.FindFrontMatterEntry("date")?.Value);
    }

    [Theory]
    [InlineData("itemize")]
    [InlineData("enumerate")]
    [InlineData("description")]
    public void EveryListKindRetainsAndDiagnosesUnrepresentedSetupAndItemlessContent(string kind) {
        string source = "\\begin{" + kind + "}\\setcounter{enumi}{4}\\item Five\\end{" + kind + "}";
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap(source)).ToMarkdownDocumentResult();
        Assert.Equal("\\setcounter{enumi}{4}", Assert.Single(result.Value.Blocks.OfType<CodeBlock>()).Content);
        Assert.Contains(result.Report.Diagnostics, static diagnostic => diagnostic.Feature == "list-content" && diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
        LatexToMarkdownResult omitted = LatexDocument.Parse(Wrap(source)).ToMarkdownDocumentResult(new LatexToMarkdownOptions { PreserveUnsupportedAsSource = false });
        Assert.Empty(omitted.Value.Blocks.OfType<CodeBlock>());
        Assert.Contains(omitted.Report.Diagnostics, static diagnostic => diagnostic.Feature == "list-content" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
        LatexToMarkdownResult itemless = LatexDocument.Parse(Wrap("\\begin{" + kind + "}IMPORTANT\\end{" + kind + "}")).ToMarkdownDocumentResult();
        Assert.Equal("IMPORTANT", Assert.Single(itemless.Value.Blocks.OfType<CodeBlock>()).Content);
        Assert.True(itemless.Report.HasLoss);
    }

    [Fact]
    public void CommonColumnAlignmentsSurviveBothDirections() {
        const string source = "\\begin{tabular}{lcr}\\textbf{L}&\\textbf{C}&\\textbf{R}\\\\left&center&right\\\\\\end{tabular}";
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap(source)).ToMarkdownDocumentResult();
        TableBlock table = Assert.Single(result.Value.Blocks.OfType<TableBlock>());
        Assert.Equal(new[] { ColumnAlignment.Left, ColumnAlignment.Center, ColumnAlignment.Right }, table.Alignments);
        Assert.False(result.Report.HasLoss);
        MarkdownToLatexResult back = result.Value.ToLatexDocumentResult();
        Assert.Equal("lcr", Assert.Single(back.Value.Tables).ColumnSpecification);
        TableBlock roundTrip = Assert.Single(back.Value.ToMarkdownDocument().Blocks.OfType<TableBlock>());
        Assert.Equal(table.Alignments, roundTrip.Alignments);
    }

    [Theory]
    [InlineData(ColumnAlignment.Center, ColumnAlignment.None, "c")]
    [InlineData(ColumnAlignment.Right, ColumnAlignment.None, "r")]
    [InlineData(ColumnAlignment.Right, ColumnAlignment.Left, "l")]
    [InlineData(ColumnAlignment.Left, ColumnAlignment.Right, "r")]
    public void SpannedCellsInheritLogicalColumnAlignmentAndRespectCellOverrides(ColumnAlignment column, ColumnAlignment cellOverride, string expected) {
        TableBlock table = Assert.Single(MarkdownReader.Parse("| H1 | H2 | H3 | H4 |\n| --- | --- | --- | --- |\n| First | Second | Third | Last |\n").Blocks.OfType<TableBlock>());
        table.Alignments[0] = ColumnAlignment.Left;
        table.Alignments[1] = ColumnAlignment.Left;
        table.Alignments[2] = column;
        TableCell first = table.GetCell(0, 0)!;
        first.ColumnSpan = 2;
        TableCell second = table.GetCell(0, 1)!;
        second.ColumnSpan = 2;
        second.Alignment = cellOverride;
        table.GetCell(0, 2)!.Alignment = ColumnAlignment.Center;
        MarkdownToLatexResult result = MarkdownDoc.Create().Add(table).ToLatexDocumentResult();
        Assert.Contains("\\multicolumn{2}{" + expected + "}{Second}", result.Source, StringComparison.Ordinal);
        Assert.Contains("\\multicolumn{1}{c}{Third}", result.Source, StringComparison.Ordinal);
        Assert.False(result.Report.HasLoss);
    }

    [Theory]
    [InlineData("|c|r|", true)]
    [InlineData("p{3cm}r", false)]
    [InlineData("*{2}{c}", false)]
    public void UnsupportedColumnLayoutSemanticsAreReportedWithoutInventingAlignment(string columns, bool aligned) {
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap("\\begin{tabular}{" + columns + "}\\textbf{C}&\\textbf{R}\\\\A&B\\\\\\end{tabular}")).ToMarkdownDocumentResult();
        TableBlock table = Assert.Single(result.Value.Blocks.OfType<TableBlock>());
        Assert.Equal(aligned ? 2 : 0, table.Alignments.Count);
        Assert.Contains(result.Report.Diagnostics, static diagnostic => diagnostic.Code == "LATEXMD215" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Simplified);
    }

    [Theory]
    [InlineData("\\resizebox{10cm}{!}{\\begin{tabular}{c}A\\\\\\end{tabular}}")]
    [InlineData("\\unknown{\\section{Title}}")]
    [InlineData("\\unknown{\\begin{itemize}\\item Item\\end{itemize}}")]
    [InlineData("\\unknown{first\n\nsecond}")]
    public void FragmentedCommandArgumentsRetainTheirCompleteFallbackAndSurroundingProse(string command) {
        LatexToMarkdownResult result = LatexDocument.Parse(Wrap("Before " + command + " After")).ToMarkdownDocumentResult();
        Assert.Equal(command, Assert.Single(result.Value.Blocks.OfType<CodeBlock>()).Content);
        Assert.Equal(2, result.Value.Blocks.OfType<ParagraphBlock>().Count());
        Assert.Empty(result.Value.Blocks.OfType<HeadingBlock>());
        Assert.Empty(result.Value.Blocks.OfType<TableBlock>());
        Assert.Empty(result.Value.Blocks.OfType<UnorderedListBlock>());
        LatexMarkdownConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, static diagnostic => diagnostic.Code == "LATEXMD297");
        Assert.StartsWith("command:", diagnostic.Feature);
        Assert.Equal(LatexMarkdownConversionOutcome.SourceFallback, diagnostic.Outcome);
        LatexToMarkdownResult omitted = LatexDocument.Parse(Wrap("Before " + command + " After")).ToMarkdownDocumentResult(new LatexToMarkdownOptions { PreserveUnsupportedAsSource = false });
        Assert.Empty(omitted.Value.Blocks.OfType<CodeBlock>());
        Assert.Equal(2, omitted.Value.Blocks.OfType<ParagraphBlock>().Count());
        Assert.Contains(omitted.Report.Diagnostics, static item => item.Outcome == LatexMarkdownConversionOutcome.Omitted);
    }

    [Fact]
    public void PreservedArgumentCommandsCannotBecomeContainerItemsCaptionsLabelsOrImages() {
        LatexToMarkdownResult list = LatexDocument.Parse(Wrap("\\begin{itemize}\\item Before \\unknown{\\item Hidden} After\\end{itemize}")).ToMarkdownDocumentResult();
        Assert.Single(Assert.Single(list.Value.Blocks.OfType<UnorderedListBlock>()).Items);
        Assert.Contains("\\unknown{\\item Hidden}", list.Value.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(list.Report.Diagnostics, static item => item.Feature == "command:unknown");
        MarkdownDoc figure = LatexDocument.Parse(Wrap("\\begin{figure}\\includegraphics{actual.png}\\unknown{\\caption{Hidden}}\\caption{Actual}\\end{figure}")).ToMarkdownDocument();
        Assert.Equal("Actual", Assert.Single(figure.Blocks.OfType<ImageBlock>()).Caption);
        Assert.Contains("\\unknown{\\caption{Hidden}}", Assert.Single(figure.Blocks.OfType<CodeBlock>()).Content, StringComparison.Ordinal);
        MarkdownDoc opaqueImage = LatexDocument.Parse(Wrap("\\begin{figure}\\phantom{\\includegraphics{hidden.png}}\\end{figure}")).ToMarkdownDocument();
        Assert.Empty(opaqueImage.Blocks.OfType<ImageBlock>());
        Assert.Contains("\\phantom{\\includegraphics{hidden.png}}", Assert.Single(opaqueImage.Blocks.OfType<CodeBlock>()).Content, StringComparison.Ordinal);
        MarkdownDoc theorem = LatexDocument.Parse(Wrap("\\begin{theorem}\\unknown{\\label{hidden}}\\label{actual}Body\\end{theorem}")).ToMarkdownDocument();
        Assert.Equal("actual", Assert.Single(theorem.Blocks.OfType<CalloutBlock>()).Attributes.ElementId);
        Assert.Contains("\\unknown{\\label{hidden}}", theorem.ToMarkdown(), StringComparison.Ordinal);
        LatexToMarkdownResult opaqueLabel = LatexDocument.Parse(Wrap("\\begin{theorem}\\unknown{A\\label{x}B}\\end{theorem}")).ToMarkdownDocumentResult();
        Assert.Null(Assert.Single(opaqueLabel.Value.Blocks.OfType<CalloutBlock>()).Attributes.ElementId);
        Assert.Contains("\\unknown{A\\label{x}B}", opaqueLabel.Value.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(opaqueLabel.Report.Diagnostics, static item => item.Feature == "command:unknown");
    }
}
