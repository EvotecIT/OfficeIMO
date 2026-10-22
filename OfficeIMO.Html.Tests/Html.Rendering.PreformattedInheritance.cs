using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPreformattedInheritanceTests {
    [Theory]
    [InlineData("span")]
    [InlineData("em")]
    [InlineData("a href='https://example.test/report'")]
    [InlineData("span><code")]
    public void DefaultPrePreservesNestedAlignmentAndLineBreaks(string openingTag) {
        string closingTag = openingTag.StartsWith("span>", StringComparison.Ordinal) ? "code></span"
            : openingTag.Split(' ')[0];
        string content = "<" + openingTag + ">Left    Right\nNext</" + closingTag + ">";
        HtmlRenderText[] actual = Render(content);
        HtmlRenderText[] explicitPre = Render(content, "white-space:pre");

        Assert.Equal("Left    RightNext", string.Concat(actual.Select(text => text.Text)));
        Assert.Equal(2, actual.Select(text => text.LayoutY).Distinct().Count());
        Assert.Equal(explicitPre.Select(TextGeometry), actual.Select(TextGeometry));
        if (openingTag.StartsWith("a ", StringComparison.Ordinal))
            Assert.All(actual, text => Assert.Equal("https://example.test/report", text.LinkUri));
    }

    [Theory]
    [InlineData("normal", false)]
    [InlineData("initial", false)]
    [InlineData("inherit", false)]
    [InlineData("unset", false)]
    [InlineData("pre", true)]
    [InlineData("pre-wrap", true)]
    [InlineData("revert", true)]
    [InlineData("revert-layer", true)]
    public void ExplicitPreValueRespectsCssResetsAndUaRollback(string value, bool preserves) {
        // An unstyled parent has CSS's normal initial whitespace value.
        HtmlRenderText[] text = Render("Left    Right\nNext", "white-space:" + value);
        Assert.Equal(preserves ? "Left    RightNext" : "Left Right Next", string.Concat(text.Select(item => item.Text)));
        Assert.Equal(preserves ? 2 : 1, text.Select(item => item.LayoutY).Distinct().Count());
    }

    [Theory]
    [InlineData("normal", false)]
    [InlineData("initial", false)]
    [InlineData("inherit", true)]
    [InlineData("unset", true)]
    public void NestedAuthorValueOverridesOrInheritsPre(string value, bool preserves) {
        HtmlRenderText[] text = Render("<span style='white-space:" + value + "'>Left    Right\nNext</span>");
        Assert.Equal(preserves ? "Left    RightNext" : "Left Right Next", string.Concat(text.Select(item => item.Text)));
        Assert.Equal(preserves ? 2 : 1, text.Select(item => item.LayoutY).Distinct().Count());
    }

    [Fact]
    public void PreDefaultOverridesImplicitParentButExplicitInheritUsesIt() {
        Assert.Equal("Left    RightNext", string.Concat(Render("<span>Left    Right\nNext</span>",
            parentStyle: "white-space:normal").Select(text => text.Text)));
        Assert.Equal("Left Right Next", string.Concat(Render("<span>Left    Right\nNext</span>",
            "white-space:inherit", "white-space:normal").Select(text => text.Text)));
    }

    [Fact]
    public void GeneratedInlineContentInheritsPreWhitespace() {
        const string content = "<span></span>";
        const string css = "pre::before{content:'Before    First\\A Second'}pre span::after{content:'After    Last'}";
        HtmlRenderText[] actual = Render(content, css: css);
        HtmlRenderText[] explicitPre = Render(content, "white-space:pre", css: css);
        Assert.Equal("Before    FirstSecondAfter    Last", string.Concat(actual.Select(text => text.Text)));
        Assert.Equal(2, actual.Select(text => text.LayoutY).Distinct().Count());
        Assert.Equal(explicitPre.Select(TextGeometry), actual.Select(TextGeometry));
    }

    private static (string Text, double X, double Y, double? Width) TextGeometry(HtmlRenderText text) =>
        (text.Text, text.X, text.LayoutY, text.TextAdvanceWidth);

    private static HtmlRenderText[] Render(string content, string style = "", string parentStyle = "", string css = "") {
        HtmlRenderDocument result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            "<style>pre{margin:0;font:12px/18px Courier}" + css + "</style><div style='" + parentStyle
            + "'><pre style='width:600px;" + style + "'>" + content + "</pre></div>"), new HtmlRenderOptions {
                Mode = HtmlRenderMode.Continuous,
                ViewportWidth = 800D,
                Margins = HtmlRenderMargins.All(0D)
            });
        return result.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
    }
}
