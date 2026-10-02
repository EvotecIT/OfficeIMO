using System.Threading;

namespace OfficeIMO.Latex.Tests;

public sealed class LatexAuditRegressionTests {
    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    [InlineData("\r")]
    public void InlineVerbatimPercentRemainsInsideItsParagraphBeforeABlankLine(string ending) {
        string source = "\\begin{document}\\verb|100%|" + ending + ending + "After\\end{document}";
        LatexDocument document = LatexDocument.Parse(source);
        Assert.Equal(new[] { "\\verb|100%|", "After" }, document.Paragraphs.Select(static paragraph => paragraph.Content));
        Assert.Equal(source, document.ToLatex());
    }

    [Theory]
    [InlineData("\\\\", new[] { "" })]
    [InlineData("&B\\\\", new[] { "", "B" })]
    [InlineData("&&C\\\\", new[] { "", "", "C" })]
    [InlineData("A&&\\\\", new[] { "A", "", "" })]
    public void TabularPreservesEmptyCellPositionsWithoutAddingATrailingRow(string content, string[] expected) {
        string source = "\\begin{tabular}{lll}" + content + "\n\\end{tabular}";
        LatexDocument document = LatexDocument.Parse(source);
        LatexTableRow row = Assert.Single(Assert.Single(document.Tables).Rows);
        Assert.Equal(expected, row.Cells.Select(static cell => cell.Content));
        Assert.Equal(Enumerable.Range(0, expected.Length), row.Cells.Select(static cell => cell.ColumnIndex));
        Assert.Equal(source, document.ToLatex());
    }

    [Fact]
    public void MacroDefinitionsDoNotContributeInactiveDocumentSemantics() {
        const string source = "\\documentclass{article}\\newcommand{\\preamble}{\\section{Inactive}}" +
            "\\begin{document}\\newcommand{\\demo}{\\section{Hidden}\\begin{tabular}{l}A\\\\\\end{tabular}\\label{fake}}Actual\\section{Visible}\\end{document}";
        LatexDocument document = LatexDocument.Parse(source);
        Assert.Equal("Visible", Assert.Single(document.Headings).Title);
        Assert.Empty(document.Tables);
        Assert.Empty(document.Labels);
        Assert.Equal(2, document.MacroDefinitions.Count);
        Assert.Contains(document.Commands, static command => command.Name == "section" && command.GetRequiredArgument(0)?.Content == "Hidden");
        Assert.Equal(source, document.ToLatex());
    }

    [Fact]
    public void MismatchedEndDoesNotClaimTerminationOrTruncateContent() {
        const string source = "\\begin{foo}text\\end{bar}trailing";
        LatexDocument document = LatexDocument.Parse(source);
        LatexEnvironment environment = Assert.Single(document.Environments);
        Assert.False(environment.IsTerminated);
        Assert.Null(environment.EndCommand);
        Assert.Equal("text\\end{bar}trailing", environment.Content);
        Assert.Equal(source, document.ToLatex());
        Assert.Contains(document.Diagnostics, static diagnostic => diagnostic.Code == "LATEX004");
    }

    [Theory]
    [InlineData("\\textbf")]
    [InlineData("\\textbf X")]
    public void UnsupportedArgumentShapeIsDiagnosedAndPreserved(string source) {
        LatexDocument document = LatexDocument.Parse(source);
        Assert.Contains(document.Diagnostics, static diagnostic => diagnostic.Code == "LATEX007");
        Assert.Equal(source, document.ToLatex());
    }

    [Fact]
    public void ConflictingEqualSpanEditsAreRejectedButIdenticalEditsAreAccepted() {
        LatexDocument document = LatexDocument.Parse("\\begin{document}original\\end{document}");
        document.Body!.Content = "container";
        Assert.Single(document.Paragraphs).Content = "child";
        Assert.Throws<InvalidOperationException>(() => document.ToLatex());
        document.Paragraphs[0].Content = "container";
        Assert.Equal("\\begin{document}container\\end{document}", document.ToLatex());
    }

    [Fact]
    public void ParserAndTokenizerEnforceInputLimitsAndPreCancellation() {
        const string source = "\n\n";
        var options = new LatexParseOptions { MaximumInputLength = 1 };
        Assert.Throws<ArgumentException>(() => LatexDocument.Parse(source, options));
        Assert.Throws<ArgumentException>(() => LatexTokenizer.Tokenize(source, options));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => LatexDocument.Parse(source, options, cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => LatexTokenizer.Tokenize(source, options, cancellation.Token));
    }

    [Fact]
    public void ExpansionLimitsAndOpaqueNamesBelongToTheParsedSnapshot() {
        var options = new LatexParseOptions { MacroExpansion = LatexMacroExpansion.SafeSimpleDefinitions, MaximumExpansionLength = 64 };
        options.VerbatimEnvironmentNames.Add("opaque");
        LatexDocument document = LatexDocument.Parse("\\newcommand{\\repeat}[1]{#1#1#1#1#1#1#1#1#1}", options);
        options.MaximumExpansionLength = 1000;
        options.VerbatimEnvironmentNames.Clear();
        Assert.Throws<System.IO.InvalidDataException>(() => document.ExpandSimpleMacros("\\repeat{long argument}"));
        Assert.Equal("\\begin{opaque}\\repeat{x}\\end{opaque}", document.ExpandSimpleMacros("\\begin{opaque}\\repeat{x}\\end{opaque}").Value);
    }
}
