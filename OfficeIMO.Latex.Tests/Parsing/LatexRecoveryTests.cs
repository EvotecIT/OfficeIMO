namespace OfficeIMO.Latex.Tests;

public sealed class LatexRecoveryTests {
    [Theory]
    [InlineData("\\textbf{broken", "LATEX001")]
    [InlineData("$broken", "LATEX003")]
    [InlineData("\\begin{document}broken", "LATEX004")]
    public void UnterminatedStructure_IsDiagnosedAndLossless(string source, string code) {
        LatexParseResult result = LatexDocument.Parse(source);

        Assert.True(result.IsLossless);
        Assert.True(result.HasErrors);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == code);
        Assert.Equal(source, result.Document.ToLatex());
    }

    [Fact]
    public void MismatchedEnvironmentEnd_IsDiagnosedWithoutDiscardingEitherCommand() {
        const string source = "\\begin{figure}x\\end{table}";

        LatexParseResult result = LatexDocument.Parse(source);

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "LATEX005");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "LATEX004");
        Assert.Equal(source, result.Document.ToLatex());
    }

    [Fact]
    public void NestingLimit_IsEnforced() {
        var options = new LatexParseOptions { MaximumNestingDepth = 3 };

        Assert.Throws<InvalidDataException>(() => LatexDocument.Parse("{{{{x}}}}", options));
    }
}
