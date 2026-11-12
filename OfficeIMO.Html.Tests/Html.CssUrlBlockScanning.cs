using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void CssUrlRewriteChecksAggregateExpansionBeforeBuildingOutput() {
        string css = string.Concat(Enumerable.Repeat("a{background:url(x)}", 32));
        Assert.Throws<InvalidDataException>(() => HtmlResourcePipeline.RewriteCssResourceUrls(css,
            (value, _) => value == "x" ? new string('y', 100) : value,
            includeFragmentReferences: false, maximumOutputCharacters: 1_000));

        string rewritten = HtmlResourcePipeline.RewriteCssResourceUrls(css,
            (value, _) => value == "x" ? "y" : value,
            includeFragmentReferences: false, maximumOutputCharacters: 1_000);
        Assert.Equal(css.Replace("url(x)", "url(\"y\")"), rewritten);
    }

    [Fact]
    public void CommentMaskingSkipsPathologicalUrlTextInsideAComment() {
        string css = "/*url(" + new string(' ', 2_048) + "*/ a {background:url(icon.png)}";
        string rewritten = HtmlResourcePipeline.RewriteCssResourceUrls(css,
            (value, _) => value == "icon.png" ? "new.png" : value);

        Assert.Equal(css.Replace("url(icon.png)", "url(\"new.png\")"), rewritten);
    }

    [Theory]
    [InlineData("url(images/*literal/icon.png)")]
    [InlineData("u\\72 l(images/*literal/icon.png)")]
    [InlineData("url(images/*literal/\\)icon.png)")]
    public void UrlPathCommentCharactersDoNotHideFollowingRuleBlocks(string url) {
        string css = "a { background: " + url + "; } @media screen { b { color: red; } }";
        var blocks = HtmlCssRuleBlockScanner.Scan(css, new HtmlCssProcessingBudget(null));
        Assert.Equal(3, blocks.Count);
        Assert.Equal(css.IndexOf('}'), blocks[css.IndexOf('{')]);
        Assert.Equal(css.LastIndexOf('}'), blocks[css.IndexOf('{', css.IndexOf("@media", StringComparison.Ordinal))]);
    }

    [Fact]
    public void UrlPathCannotHideNestingBudgetViolationsInFollowingRules() {
        const string css = "a { background:url(images/*literal/icon.png); } @media screen { @supports (color:red) { b {color:red} } }";
        Assert.Throws<HtmlDomLimitException>(() => HtmlResourcePipeline.BuildManifest("<style>" + css + "</style>",
            new HtmlResourcePipelineOptions { Limits = new HtmlConversionLimits { MaxCssNestingDepth = 2 } }));
    }

    [Theory]
    [InlineData("url/**/( { } )")]
    [InlineData("u/**/rl( { } )")]
    [InlineData("url(icon.png) /* { ignored } */ { }")]
    public void CommentsOutsideUrlTokensKeepTheirTokenBoundary(string css) {
        var blocks = HtmlCssRuleBlockScanner.Scan(css, new HtmlCssProcessingBudget(null));
        Assert.Single(blocks);
        Assert.Equal(css.LastIndexOf('}'), blocks.Values.Single());
    }

    [Theory]
    [InlineData("url(\"icon.png\"/* ) {{{ */)")]
    [InlineData("url( 'icon.png' /* ) {{{ */)")]
    public void QuotedUrlFunctionsRetainCommentsAfterTheirString(string url) {
        string css = "a {background:" + url + ";} b {color:red}";
        var blocks = HtmlCssRuleBlockScanner.Scan(css,
            new HtmlCssProcessingBudget(new HtmlConversionLimits { MaxCssNestingDepth = 1 }));
        Assert.Equal(2, blocks.Count);
        Assert.Equal(css.IndexOf(";} b", StringComparison.Ordinal) + 1, blocks[css.IndexOf('{')]);
    }

    [Theory]
    [InlineData("@font-face { }", true)]
    [InlineData("@font-face/**/{ }", true)]
    [InlineData("@font-face /* a */ /* b */ { }", true)]
    [InlineData(".x\\@font-face { }", false)]
    [InlineData("@font-face */ { }", false)]
    public void FontFaceBlocksAreRecognizedOnlyFromTheAtRuleToken(string css, bool expected) {
        Assert.Equal(expected, HtmlCssRuleBlockScanner.IsFontFaceBlock(css, css.IndexOf('{')));
    }

    [Fact]
    public void NestedMalformedFontFaceBlocksDoNotStopTheCascade() {
        string payload = new string('A', 2_048);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@font-face{font-family:A;@font-face{src:url(data:font/ttf;base64," + payload + ")}}"
            + " p{color:blue}</style><p>Text</p>");

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("p")!];

        Assert.Equal("rgba(0, 0, 255, 1)", style.GetValue("color"));
    }
}
