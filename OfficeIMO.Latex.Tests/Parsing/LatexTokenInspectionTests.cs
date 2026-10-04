using System.Collections;
using System.Threading.Tasks;

namespace OfficeIMO.Latex.Tests;

public sealed class LatexTokenInspectionTests {
    [Fact]
    public void SparseAndCompleteInspectionKeepExactSpansValuesAndStableIdentity() {
        string source = string.Concat(Enumerable.Repeat("\\textbf A % comment\r\n$y_1$ ", 1100)) + "\\verb|unfinished";
        LatexDocument document = LatexDocument.Parse(source);
        IReadOnlyList<LatexToken> tokens = document.Tokens;
        LatexToken middle = tokens[tokens.Count / 2];
        LatexToken[] complete = tokens.ToArray();
        Assert.Same(middle, complete[tokens.Count / 2]);
        Assert.Equal(source, string.Concat(complete.Select(token => token.Text)));
        int offset = 0;
        for (int index = 0; index < complete.Length; index++) {
            LatexToken token = complete[index];
            Assert.Same(token, tokens[index]);
            Assert.Equal(offset, token.Span.Start.Offset);
            Assert.Equal(source.Substring(offset, token.Span.Length), token.Text);
            offset = token.Span.End.Offset;
            if (token.Kind == LatexTokenKind.Command) Assert.Equal("textbf", token.Value);
            else if (token.Kind != LatexTokenKind.Verbatim) Assert.Null(token.Value);
        }
        Assert.Equal(source.Length, offset);
        Assert.False(complete[complete.Length - 1].IsTerminated);
        Assert.Equal("verb", complete[complete.Length - 1].Value);
        Assert.Equal(complete, ((IEnumerable)tokens).Cast<LatexToken>().ToArray());
        Assert.Equal(source, document.ToLatex());
        Assert.Throws<ArgumentOutOfRangeException>(() => tokens[-1]);
        Assert.Throws<ArgumentOutOfRangeException>(() => tokens[tokens.Count]);
    }

    [Fact]
    public void ConcurrentInspectionCannotChangeCommandValuesOrTokenIdentity() {
        string source = string.Concat(Enumerable.Repeat("\\section{Title} \\verb|x|\n", 256));
        IReadOnlyList<LatexToken> tokens = LatexTokenizer.Tokenize(source);
        var observed = new LatexToken[tokens.Count];
        Parallel.For(0, 16, iteration => {
            for (int index = 0; index < tokens.Count; index++) {
                LatexToken token = tokens[index];
                LatexToken? previous = System.Threading.Interlocked.CompareExchange(ref observed[index], token, null);
                if (previous != null) Assert.Same(previous, token);
                Assert.Equal(source.Substring(token.Span.Start.Offset, token.Span.Length), token.Text);
                if (token.Kind == LatexTokenKind.Command) Assert.Equal("section", token.Value);
                if (token.Kind == LatexTokenKind.Verbatim) Assert.Equal("verb", token.Value);
            }
        });
        Assert.Equal(source, string.Concat(tokens.Select(token => token.Text)));
    }

    [Fact]
    public void EmptyAndSingleTokenSourcesRemainInspectable() {
        Assert.Empty(LatexTokenizer.Tokenize(string.Empty));
        Assert.Equal("x", Assert.Single(LatexTokenizer.Tokenize("x")).Text);
        Assert.Equal("\\", Assert.Single(LatexTokenizer.Tokenize("\\")).Text);
    }
}
