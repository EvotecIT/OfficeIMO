using OfficeIMO.Pdf;

namespace OfficeIMO.Latex.Markdown.Tests;

public sealed class LatexForwardOwnerTests {
    private static string Wrap(string body) => "\\documentclass{article}\\begin{document}" + body + "\\end{document}";

    [Fact]
    public void PreservedPartsCannotShiftVisibleHeadingLevels() {
        MarkdownDocument result = LatexDocument.Parse(Wrap("\\unknown{\\part{Hidden}}\\section{Visible}")).ToMarkdown();
        Assert.Equal(1, Assert.Single(result.Blocks.OfType<HeadingBlock>()).Level);
        Assert.Contains("\\unknown{\\part{Hidden}}", Assert.Single(result.Blocks.OfType<CodeBlock>()).Content, StringComparison.Ordinal);
    }

    [Fact]
    public void InertMetadataCannotShadowActivePreambleOrCreateProjectedMetadata() {
        const string trailer = "\\title{INERT TITLE}\\author{PRIVATE}\\date{INERT DATE}\\documentclass{report}";
        var result = LatexDocument.Parse(Wrap("Public") + trailer).ToMarkdown(null);
        Assert.Null(result.FindFrontMatterEntry("title"));
        Assert.Null(result.FindFrontMatterEntry("author"));
        Assert.Null(result.FindFrontMatterEntry("date"));
        Assert.Equal("article", result.FindFrontMatterEntry("documentclass")?.Value);
        Assert.Empty(result.Blocks.OfType<HeadingBlock>());
        string text = PdfDocument.Load(new System.IO.MemoryStream(LatexDocument.Parse(Wrap("Public") + trailer).ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }))).ExtractText();
        Assert.Contains("Public", text, StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE", text, StringComparison.Ordinal);
        Assert.DoesNotContain("INERT", text, StringComparison.Ordinal);
        MarkdownDocument active = LatexDocument.Parse("\\title{Active}\\author{Author}\\date{Today}" + Wrap("Body") + trailer).ToMarkdown();
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
        var result = LatexDocument.Parse(Wrap(source)).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport2);
        var resultReport = (LatexToMarkdownReport)conversionReport2;
        Assert.Equal("\\setcounter{enumi}{4}", Assert.Single(result.Blocks.OfType<CodeBlock>()).Content);
        Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Feature == "list-content" && diagnostic.Outcome == LatexMarkdownConversionOutcome.SourceFallback);
        var omitted = LatexDocument.Parse(Wrap(source)).ToMarkdown(new LatexMarkdownExportOptions { PreserveUnsupportedAsSource = false }, out OfficeIMO.OfficeConversionReport conversionReport3);
        var omittedReport = (LatexToMarkdownReport)conversionReport3;
        Assert.Empty(omitted.Blocks.OfType<CodeBlock>());
        Assert.Contains(omittedReport.NativeDiagnostics, static diagnostic => diagnostic.Feature == "list-content" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Omitted);
        var itemless = LatexDocument.Parse(Wrap("\\begin{" + kind + "}IMPORTANT\\end{" + kind + "}")).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport4);
        var itemlessReport = (LatexToMarkdownReport)conversionReport4;
        Assert.Equal("IMPORTANT", Assert.Single(itemless.Blocks.OfType<CodeBlock>()).Content);
        Assert.True(itemlessReport.HasLoss);
    }

    [Fact]
    public void CommonColumnAlignmentsSurviveBothDirections() {
        const string source = "\\begin{tabular}{lcr}\\textbf{L}&\\textbf{C}&\\textbf{R}\\\\left&center&right\\\\\\end{tabular}";
        var result = LatexDocument.Parse(Wrap(source)).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport5);
        var resultReport = (LatexToMarkdownReport)conversionReport5;
        TableBlock table = Assert.Single(result.Blocks.OfType<TableBlock>());
        Assert.Equal(new[] { MarkdownColumnAlignment.Left, MarkdownColumnAlignment.Center, MarkdownColumnAlignment.Right }, table.Alignments);
        Assert.False(resultReport.HasLoss);
        var back = result.ToLatex(null);
        Assert.Equal("lcr", Assert.Single(back.Tables).ColumnSpecification);
        TableBlock roundTrip = Assert.Single(back.ToMarkdown().Blocks.OfType<TableBlock>());
        Assert.Equal(table.Alignments, roundTrip.Alignments);
    }

    [Fact]
    public void SparseTableAlignmentCannotExpandPastProjectedCellBudget() {
        string rows = string.Concat(Enumerable.Repeat("A\\\\", 256));
        string source = Wrap("\\begin{tabular}{" + new string('l', 1024) + "}" + rows + "\\end{tabular}");
        Assert.Throws<System.IO.InvalidDataException>(() => LatexDocument.Parse(source).ToMarkdown());
    }

    [Fact]
    public void ColumnAlignmentCannotExceedProjectedColumnBudget() {
        string source = Wrap("\\begin{tabular}{" + new string('l', 4097) + "}A\\\\\\end{tabular}");
        Assert.Throws<System.IO.InvalidDataException>(() => LatexDocument.Parse(source).ToMarkdown());
    }

    [Fact]
    public void MultipleSparseTablesShareProjectionCellBudget() {
        string table = "\\begin{tabular}{" + new string('l', 128) + "}"
            + string.Concat(Enumerable.Repeat("A\\\\", 256)) + "\\end{tabular}";
        Assert.Throws<System.IO.InvalidDataException>(() => LatexDocument.Parse(Wrap(table + table)).ToMarkdown());
    }

    [Theory]
    [InlineData(MarkdownColumnAlignment.Center, MarkdownColumnAlignment.None, "c")]
    [InlineData(MarkdownColumnAlignment.Right, MarkdownColumnAlignment.None, "r")]
    [InlineData(MarkdownColumnAlignment.Right, MarkdownColumnAlignment.Left, "l")]
    [InlineData(MarkdownColumnAlignment.Left, MarkdownColumnAlignment.Right, "r")]
    public void SpannedCellsInheritLogicalColumnAlignmentAndRespectCellOverrides(MarkdownColumnAlignment column, MarkdownColumnAlignment cellOverride, string expected) {
        TableBlock table = Assert.Single(MarkdownDocument.Parse("| H1 | H2 | H3 | H4 |\n| --- | --- | --- | --- |\n| First | Second | Third | Last |\n").Blocks.OfType<TableBlock>());
        table.Alignments[0] = MarkdownColumnAlignment.Left;
        table.Alignments[1] = MarkdownColumnAlignment.Left;
        table.Alignments[2] = column;
        TableCell first = table.GetCell(0, 0)!;
        first.ColumnSpan = 2;
        TableCell second = table.GetCell(0, 1)!;
        second.ColumnSpan = 2;
        second.Alignment = cellOverride;
        table.GetCell(0, 2)!.Alignment = MarkdownColumnAlignment.Center;
        using var addedDocument = MarkdownDocument.Create();
        addedDocument.Add(table);
        var result = addedDocument.ToLatex(null, out OfficeIMO.OfficeConversionReport conversionReport10);
        var resultReport = (MarkdownToLatexReport)conversionReport10;
        Assert.Contains("\\multicolumn{2}{" + expected + "}{Second}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.Contains("\\multicolumn{1}{c}{Third}", resultReport.GeneratedSource, StringComparison.Ordinal);
        Assert.False(resultReport.HasLoss);
    }

    [Theory]
    [InlineData("|c|r|", true)]
    [InlineData("p{3cm}r", false)]
    [InlineData("*{2}{c}", false)]
    public void UnsupportedColumnLayoutSemanticsAreReportedWithoutInventingAlignment(string columns, bool aligned) {
        var result = LatexDocument.Parse(Wrap("\\begin{tabular}{" + columns + "}\\textbf{C}&\\textbf{R}\\\\A&B\\\\\\end{tabular}")).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport11);
        var resultReport = (LatexToMarkdownReport)conversionReport11;
        TableBlock table = Assert.Single(result.Blocks.OfType<TableBlock>());
        Assert.Equal(aligned ? 2 : 0, table.Alignments.Count);
        Assert.Contains(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Code == "LATEXMD215" && diagnostic.Outcome == LatexMarkdownConversionOutcome.Simplified);
    }

    [Theory]
    [InlineData("\\resizebox{10cm}{!}{\\begin{tabular}{c}A\\\\\\end{tabular}}")]
    [InlineData("\\unknown{\\section{Title}}")]
    [InlineData("\\unknown{\\begin{itemize}\\item Item\\end{itemize}}")]
    [InlineData("\\unknown{first\n\nsecond}")]
    public void FragmentedCommandArgumentsRetainTheirCompleteFallbackAndSurroundingProse(string command) {
        var result = LatexDocument.Parse(Wrap("Before " + command + " After")).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport12);
        var resultReport = (LatexToMarkdownReport)conversionReport12;
        Assert.Equal(command, Assert.Single(result.Blocks.OfType<CodeBlock>()).Content);
        Assert.Equal(2, result.Blocks.OfType<ParagraphBlock>().Count());
        Assert.Empty(result.Blocks.OfType<HeadingBlock>());
        Assert.Empty(result.Blocks.OfType<TableBlock>());
        Assert.Empty(result.Blocks.OfType<UnorderedListBlock>());
        LatexMarkdownConversionDiagnostic diagnostic = Assert.Single(resultReport.NativeDiagnostics, static diagnostic => diagnostic.Code == "LATEXMD297");
        Assert.StartsWith("command:", diagnostic.Feature);
        Assert.Equal(LatexMarkdownConversionOutcome.SourceFallback, diagnostic.Outcome);
        var omitted = LatexDocument.Parse(Wrap("Before " + command + " After")).ToMarkdown(new LatexMarkdownExportOptions { PreserveUnsupportedAsSource = false }, out OfficeIMO.OfficeConversionReport conversionReport13);
        var omittedReport = (LatexToMarkdownReport)conversionReport13;
        Assert.Empty(omitted.Blocks.OfType<CodeBlock>());
        Assert.Equal(2, omitted.Blocks.OfType<ParagraphBlock>().Count());
        Assert.Contains(omittedReport.NativeDiagnostics, static item => item.Outcome == LatexMarkdownConversionOutcome.Omitted);
    }

    [Fact]
    public void PreservedArgumentCommandsCannotBecomeContainerItemsCaptionsLabelsOrImages() {
        var list = LatexDocument.Parse(Wrap("\\begin{itemize}\\item Before \\unknown{\\item Hidden} After\\end{itemize}")).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport14);
        var listReport = (LatexToMarkdownReport)conversionReport14;
        Assert.Single(Assert.Single(list.Blocks.OfType<UnorderedListBlock>()).Items);
        Assert.Contains("\\unknown{\\item Hidden}", list.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(listReport.NativeDiagnostics, static item => item.Feature == "command:unknown");
        MarkdownDocument figure = LatexDocument.Parse(Wrap("\\begin{figure}\\includegraphics{actual.png}\\unknown{\\caption{Hidden}}\\caption{Actual}\\end{figure}")).ToMarkdown();
        Assert.Equal("Actual", Assert.Single(figure.Blocks.OfType<ImageBlock>()).Caption);
        Assert.Contains("\\unknown{\\caption{Hidden}}", Assert.Single(figure.Blocks.OfType<CodeBlock>()).Content, StringComparison.Ordinal);
        MarkdownDocument opaqueImage = LatexDocument.Parse(Wrap("\\begin{figure}\\phantom{\\includegraphics{hidden.png}}\\end{figure}")).ToMarkdown();
        Assert.Empty(opaqueImage.Blocks.OfType<ImageBlock>());
        Assert.Contains("\\phantom{\\includegraphics{hidden.png}}", Assert.Single(opaqueImage.Blocks.OfType<CodeBlock>()).Content, StringComparison.Ordinal);
        MarkdownDocument theorem = LatexDocument.Parse(Wrap("\\begin{theorem}\\unknown{\\label{hidden}}\\label{actual}Body\\end{theorem}")).ToMarkdown();
        Assert.Equal("actual", Assert.Single(theorem.Blocks.OfType<CalloutBlock>()).Attributes.ElementId);
        Assert.Contains("\\unknown{\\label{hidden}}", theorem.ToMarkdown(), StringComparison.Ordinal);
        var opaqueLabel = LatexDocument.Parse(Wrap("\\begin{theorem}\\unknown{A\\label{x}B}\\end{theorem}")).ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport15);
        var opaqueLabelReport = (LatexToMarkdownReport)conversionReport15;
        Assert.Null(Assert.Single(opaqueLabel.Blocks.OfType<CalloutBlock>()).Attributes.ElementId);
        Assert.Contains("\\unknown{A\\label{x}B}", opaqueLabel.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains(opaqueLabelReport.NativeDiagnostics, static item => item.Feature == "command:unknown");
    }
}
