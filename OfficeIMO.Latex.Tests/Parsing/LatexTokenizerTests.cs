namespace OfficeIMO.Latex.Tests;

public sealed class LatexTokenizerTests {
    [Fact]
    public void Tokenizer_CoversCommandsCommentsEscapesMathAndMixedLineEndingsExactly() {
        const string source = "\\section{A} % comment\r\nText \\% $x_1$ $$y^2$$\r";

        IReadOnlyList<LatexToken> tokens = LatexTokenizer.Tokenize(source);

        Assert.Equal(source, string.Concat(tokens.Select(static token => token.Text)));
        Assert.Contains(tokens, token => token.Kind == LatexTokenKind.Command && token.Value == "section");
        Assert.Contains(tokens, token => token.Kind == LatexTokenKind.Command && token.Value == "%");
        Assert.Contains(tokens, token => token.Kind == LatexTokenKind.Comment && token.Text == "% comment");
        Assert.Equal(4, tokens.Count(token => token.Kind == LatexTokenKind.MathShift));
        Assert.Contains(tokens, token => token.Kind == LatexTokenKind.Subscript);
        Assert.Contains(tokens, token => token.Kind == LatexTokenKind.Superscript);
        Assert.Equal(new[] { "\r\n", "\r" }, tokens.Where(token => token.Kind == LatexTokenKind.LineEnding).Select(token => token.Text));
    }

    [Fact]
    public void TokenLimit_IsEnforcedWithoutPartialResult() {
        var options = new LatexParseOptions { MaximumTokenCount = 2 };

        Assert.Throws<InvalidDataException>(() => LatexTokenizer.Tokenize("a b c", options));
    }
}
