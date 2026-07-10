namespace OfficeIMO.Latex.Tests;

public sealed class LatexMacroSafetyRegressionTests {
    [Fact]
    public void RenewCommand_ReplacesEarlierDocumentLocalDefinitionInSourceOrder() {
        const string source = "\\newcommand{\\term}{old}\\renewcommand{\\term}{new}";
        LatexDocument document = LatexDocument.Parse(source, new LatexParseOptions {
            MacroExpansion = LatexMacroExpansion.SafeSimpleDefinitions
        }).Document;

        LatexMacroExpansionResult result = document.ExpandSimpleMacros("\\term");

        Assert.Equal("new", result.Value);
        Assert.Empty(result.Diagnostics);
    }

    [Theory]
    [InlineData("\\edef\\x{value}")]
    [InlineData("\\gdef\\x{value}")]
    [InlineData("\\let\\x\\y")]
    [InlineData("\\usepackage{shellesc}")]
    [InlineData("\\includegraphics{secret}")]
    public void UnsafeOrSideEffectingControlSequences_AreNotClassifiedAsSafe(string body) {
        string source = "\\newcommand{\\candidate}{" + body + "}";

        LatexMacroDefinition definition = Assert.Single(LatexDocument.Parse(source).Document.MacroDefinitions);

        Assert.False(definition.IsSafe);
    }

    [Fact]
    public void SafeClassification_IsTransitiveAcrossDocumentLocalMacros() {
        const string source =
            "\\documentclass{article}\n" +
            "\\newcommand{\\unsafe}{\\input{secret.tex}}\n" +
            "\\newcommand{\\wrapper}{\\unsafe}\n" +
            "\\begin{document}\\wrapper\\end{document}\n";

        LatexDocument document = LatexDocument.Parse(source, new LatexParseOptions {
            MacroExpansion = LatexMacroExpansion.SafeSimpleDefinitions
        }).Document;

        Assert.All(document.MacroDefinitions, static definition => Assert.False(definition.IsSafe));
        LatexMacroExpansionResult expansion = document.ExpandSimpleMacros("\\wrapper");
        Assert.Equal("\\wrapper", expansion.Value);
    }

    [Theory]
    [InlineData("\\newcommand{\\candidate}[many]{value}")]
    [InlineData("\\newcommand{\\candidate}[0][default]{value}")]
    public void MalformedSimpleDefinitionOptions_AreNotClassifiedAsSafe(string source) {
        LatexMacroDefinition definition = Assert.Single(LatexDocument.Parse(source).Document.MacroDefinitions);

        Assert.False(definition.IsSafe);
    }
}
