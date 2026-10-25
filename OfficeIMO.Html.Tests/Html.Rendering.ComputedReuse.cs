using AngleSharp.Dom;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlComputedStyleReuseTests {
    [Fact]
    public void RepeatedSiblingsKeepTheirInheritedFontSizeAndCustomProperties() {
        const string html = "<style>"
            + ".parent{font-size:20px;font-weight:400;--tone:red}"
            + ".alternate{font-size:40px;font-weight:700;--tone:blue}"
            + ".child{font-size:50%;font-weight:bolder;color:var(--tone)}"
            + "</style><div class='parent'><span class='child' id='first'>A</span>"
            + "<span class='child' id='second'>B</span></div>"
            + "<div class='alternate'><span class='child' id='third'>C</span>"
            + "<span class='child' id='fourth'>D</span></div>";
        var (document, styles) = Compute(html);

        foreach (string id in new[] { "first", "second" }) {
            HtmlComputedStyle style = styles.Elements[document.QuerySelector("#" + id)!];
            Assert.Equal(7.5D, style.ResolvedFontSizePoints ?? 0D, 5);
            Assert.Equal("700", style.GetValue("font-weight"));
            Assert.Equal("red", style.GetValue("color"));
        }
        foreach (string id in new[] { "third", "fourth" }) {
            HtmlComputedStyle style = styles.Elements[document.QuerySelector("#" + id)!];
            Assert.Equal(15D, style.ResolvedFontSizePoints ?? 0D, 5);
            Assert.Equal("900", style.GetValue("font-weight"));
            Assert.Equal("blue", style.GetValue("color"));
        }
    }

    [Fact]
    public void EquivalentAuthorDeclarationsKeepHiddenDialogAndMathDisplayDefaults() {
        var (document, styles) = Compute("<div id='visible'>A</div><div hidden id='hidden'>B</div>"
            + "<dialog id='closed'>C</dialog><dialog open id='open'>D</dialog>"
            + "<math display='block' id='blockMath'><mi>x</mi></math>"
            + "<math id='inlineMath'><mi>y</mi></math>");
        AssertDisplay("visible", "block");
        AssertDisplay("hidden", "none");
        AssertDisplay("closed", "none");
        AssertDisplay("open", "block");
        AssertDisplay("blockMath", "block");
        AssertDisplay("inlineMath", "inline");

        void AssertDisplay(string id, string expected) =>
            Assert.Equal(expected, styles.Elements[document.QuerySelector("#" + id)!].GetValue("display"));
    }

    [Fact]
    public void RepeatedInlineValuesKeepTheirPropertySpecificGrammar() {
        var (document, styles) = Compute("<div id='first' style='border-style:solid;border-width:solid'></div>"
            + "<div id='second' style='border-width:solid;border-style:solid'></div>");
        foreach (string id in new[] { "first", "second" }) {
            HtmlComputedStyle style = styles.Elements[document.QuerySelector("#" + id)!];
            Assert.Equal("solid", style.GetValue("border-top-style"));
            Assert.NotEqual("solid", style.GetValue("border-top-width"));
        }
    }

    private static (AngleSharp.Html.Dom.IHtmlDocument Document, HtmlComputedStyleSet Styles) Compute(string html) {
        var document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();
        return (document, HtmlComputedStyleEngine.ComputeForRendering(document, new HtmlRenderOptions(),
            HtmlConversionLimits.CreateUntrustedProfile()));
    }
}
