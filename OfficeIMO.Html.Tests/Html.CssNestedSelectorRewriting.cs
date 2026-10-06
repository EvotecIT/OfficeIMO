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
    [Fact]
    public void ExpandedAttributeAlternativesPreserveMatchingAndClassSpecificity() {
        string css = HtmlCssIdSelectorRewriter.Rewrite("[name=old] {font-weight:700} .later {font-weight:400}",
            new Dictionary<string, string>(), CancellationToken.None,
            (_, _, _) => HtmlCssAttributeSelectorEdit.Exact(new[] { "first", "second" }));
        var document = HtmlDocumentParser.ParseDocument("<style>" + css + "</style><p name='first'>First</p><p name='second' class='later'>Second</p><p name='other'>Other</p>");
        var styles = HtmlComputedStyleEngine.Compute(document);
        var paragraphs = document.QuerySelectorAll("p");
        Assert.Equal("700", styles[paragraphs[0]].GetValue("font-weight"));
        Assert.Equal("400", styles[paragraphs[1]].GetValue("font-weight"));
        Assert.NotEqual("700", styles[paragraphs[2]].GetValue("font-weight"));
    }
}
