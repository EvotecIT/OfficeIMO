using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Theory]
    [InlineData("url(images/*literal/icon.png)")]
    [InlineData("\\75 rl(images/*literal/\\)icon.png)")]
    [InlineData("url(\"icon.png\"/* ) {{{ */)")]
    [InlineData("fn([url(images/*literal/icon.png)], {value:'#heading'})")]
    public void NestedSelectorRewritePreservesUrlTokensInsideDeclarationValues(string value) {
        string css = ".chapter { --asset:" + value + "; & #heading {color:red;} }";
        string result = HtmlCssIdSelectorRewriter.Rewrite(css,
            new Dictionary<string, string> { ["heading"] = "second-heading" }, CancellationToken.None);
        Assert.Equal(css.Replace("& #heading", "& #second-heading"), result);
    }
}
