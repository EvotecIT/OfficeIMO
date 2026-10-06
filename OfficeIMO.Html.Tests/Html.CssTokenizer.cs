using System.Threading;
using OfficeIMO.Html.Css;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlCssTokenizerTests {
    [Theory]
    [InlineData("co\\6c or", HtmlCssTokenKind.Identifier, "color")]
    [InlineData("--\\1f600 ", HtmlCssTokenKind.Identifier, "--😀")]
    [InlineData("@m\\65 dia", HtmlCssTokenKind.AtKeyword, "media")]
    [InlineData("1.2e-3p\\78", HtmlCssTokenKind.Dimension, "px")]
    [InlineData("+1e3", HtmlCssTokenKind.Number, null)]
    [InlineData("-.5%", HtmlCssTokenKind.Percentage, null)]
    [InlineData("f\\6f o(", HtmlCssTokenKind.Function, "foo")]
    [InlineData("url(data:a;b{c})", HtmlCssTokenKind.Url, "data:a;b{c}")]
    [InlineData("u\\72l(a\\)b)", HtmlCssTokenKind.Url, "a)b")]
    [InlineData("url(a b;{x})", HtmlCssTokenKind.BadUrl, null)]
    [InlineData("url(a\\\nb;c)", HtmlCssTokenKind.BadUrl, null)]
    [InlineData("'a\\\r\nb'", HtmlCssTokenKind.String, "ab")]
    [InlineData("'a\\0 b'", HtmlCssTokenKind.String, "a�b")]
    [InlineData("'unterminated", HtmlCssTokenKind.String, "unterminated")]
    [InlineData("\\", HtmlCssTokenKind.Identifier, "�")]
    [InlineData("\0x", HtmlCssTokenKind.Identifier, "�x")]
    [InlineData("/* unfinished", HtmlCssTokenKind.Comment, null)]
    [InlineData("<!--", HtmlCssTokenKind.Cdo, null)]
    [InlineData("-->", HtmlCssTokenKind.Cdc, null)]
    public void DecodesLexicalValuesWithoutLosingAuthoredText(string css, HtmlCssTokenKind kind, string? value) {
        var tokens = HtmlCssTokenizer.Tokenize(css);
        HtmlCssToken token = Assert.Single(tokens, item => item.Kind != HtmlCssTokenKind.EndOfFile);
        Assert.Equal(kind, token.Kind);
        Assert.Equal(value, token.Value);
        Assert.Equal(css, token.GetText(css));
        Assert.Equal(css.Length, tokens[tokens.Count - 1].Offset);
    }

    [Fact]
    public void RetainsCommentsWithoutMergingIdentifiersAndRecognizesQuotedUrls() {
        const string css = "a/**/b url( \"a;b\" ) #12 #name";
        var tokens = HtmlCssTokenizer.Tokenize(css);
        Assert.Equal(new[] { HtmlCssTokenKind.Identifier, HtmlCssTokenKind.Comment, HtmlCssTokenKind.Identifier }, tokens.Take(3).Select(token => token.Kind));
        Assert.Equal("url", Assert.Single(tokens, token => token.Kind == HtmlCssTokenKind.Function).Value);
        Assert.Equal("a;b", Assert.Single(tokens, token => token.Kind == HtmlCssTokenKind.String).Value);
        Assert.Equal(new[] { false, true }, tokens.Where(token => token.Kind == HtmlCssTokenKind.Hash).Select(token => token.IsIdentifierHash));
        Assert.Equal(css, string.Concat(tokens.Select(token => token.GetText(css))));
    }

    [Fact]
    public void RecoversBadStringsAtTheOriginalNewlineAndPreservesSourceOffsets() {
        const string css = "'bad\r\n;color:blue";
        var tokens = HtmlCssTokenizer.Tokenize(css);
        Assert.Equal(HtmlCssTokenKind.BadString, tokens[0].Kind);
        Assert.Equal("'bad", tokens[0].GetText(css));
        Assert.Equal(HtmlCssTokenKind.Whitespace, tokens[1].Kind);
        Assert.Equal("\r\n", tokens[1].GetText(css));
        Assert.Equal(6, tokens[2].Offset);
        Assert.Equal(HtmlCssTokenKind.Semicolon, tokens[2].Kind);
        Assert.Equal(css, string.Concat(tokens.Select(token => token.GetText(css))));
    }

    [Fact]
    public void ReplacesUnpairedSurrogatesButPreservesNonBmpCodePoints() {
        // Construct invalid UTF-16 at runtime: custom-attribute string encoding cannot preserve it.
        string source = new string(new[] { (char)0xd800, 'x', (char)0xdc00 });
        Assert.Equal("�x�", HtmlCssTokenizer.Tokenize(source)[0].Value);
        Assert.Equal(source, HtmlCssTokenizer.Tokenize(source)[0].GetText(source));
        Assert.Equal("😀", HtmlCssTokenizer.Tokenize("\\😀")[0].Value);
    }

    [Fact]
    public void EnforcesStandaloneLimitsAndCancellationWithoutPartialResults() {
        Assert.Equal("MaxInputCharacters", Assert.Throws<HtmlCssTokenizationLimitException>(() => HtmlCssTokenizer.Tokenize("abc", new HtmlCssTokenizationOptions { MaxInputCharacters = 2 })).LimitName);
        var failure = Assert.Throws<HtmlCssTokenizationLimitException>(() => HtmlCssTokenizer.Tokenize("a:b", new HtmlCssTokenizationOptions { MaxTokens = 2 }));
        Assert.Equal("MaxTokens", failure.LimitName);
        Assert.Equal(3, failure.Actual);
        Assert.Equal(4, HtmlCssTokenizer.Tokenize("a:b", new HtmlCssTokenizationOptions { MaxTokens = 3 }).Count);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => HtmlCssTokenizer.Tokenize("/* text */", cancellationToken: cancellation.Token));
        Assert.Throws<ArgumentOutOfRangeException>(() => HtmlCssTokenizer.Tokenize("", new HtmlCssTokenizationOptions { MaxTokens = 0 }));
    }

    [Fact]
    public void MalformedInputMakesProgressAndRetainsEverySourceCharacter() {
        const string alphabet = "ab0-+\\\r\n\0\ud800\udc00()[]{}:;/*'\"#@!<>. ";
        var random = new Random(4812);
        for (int sample = 0; sample < 200; sample++) {
            string source = new string(Enumerable.Range(0, 150).Select(_ => alphabet[random.Next(alphabet.Length)]).ToArray());
            var tokens = HtmlCssTokenizer.Tokenize(source);
            int position = 0;
            foreach (HtmlCssToken token in tokens) {
                Assert.Equal(position, token.Offset);
                if (token.Kind != HtmlCssTokenKind.EndOfFile) Assert.True(token.Length > 0);
                position += token.Length;
            }
            Assert.Equal(source.Length, position);
            Assert.Equal(source, string.Concat(tokens.Select(token => token.GetText(source))));
        }
    }
}
