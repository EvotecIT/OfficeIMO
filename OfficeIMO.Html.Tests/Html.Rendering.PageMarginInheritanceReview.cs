using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderUserAgentStyleMode.Document, false, "font-family:initial")]
    [InlineData(HtmlRenderUserAgentStyleMode.Document, true, "font-family:initial")]
    [InlineData(HtmlRenderUserAgentStyleMode.Document, false, "font:initial")]
    [InlineData(HtmlRenderUserAgentStyleMode.Document, true, "font:initial")]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser, false, "font-family:initial")]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser, true, "font:initial")]
    public void HtmlPageMarginInheritance_InitialFamilyUsesTheConfiguredInitialRatherThanItsParent(
        HtmlRenderUserAgentStyleMode mode, bool resetPage, string declaration) {
        HtmlRenderOptions options = PageMarginTypographyOptions();
        options.UserAgentStyles = mode;
        options.DefaultFontFamily = "Pinned";
        string html = PageMarginInheritanceSource("font:20px/1.5 Arial",
            resetPage ? declaration : "font-family:Arial", resetPage ? "" : declaration);

        HtmlRenderText margin = PageMarginTypographyText(Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), options).Pages));

        Assert.Equal("Pinned", margin.Font.FamilyName);
    }

    [Theory]
    [InlineData("", "", 75D, 500)]
    [InlineData("font-stretch:normal", "", 100D, 500)]
    [InlineData("", "font-stretch:normal", 100D, 500)]
    [InlineData("font-stretch:expanded", "font-stretch:inherit", 125D, 500)]
    [InlineData("", "font:initial", 100D, 400)]
    [InlineData("font:initial", "", 100D, 400)]
    [InlineData("font-weight:900", "", 75D, 900)]
    public void HtmlPageMarginInheritance_DescriptorOverridesAndResetsReachThePublicPaintScene(
        string pageDeclaration, string marginDeclaration, double expectedStretch, int expectedWeight) {
        string html = PageMarginInheritanceSource("font:20px/1.5 Pinned;font-stretch:75%;font-weight:500",
            pageDeclaration, marginDeclaration);
        HtmlRenderText margin = PageMarginTypographyText(Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages));

        Assert.Equal(expectedStretch, margin.Font.Face.StretchPercent, 6);
        Assert.Equal(expectedStretch, margin.FontDescriptor.StretchPercent, 6);
        Assert.Equal(expectedWeight, margin.Font.Face.Weight);
        Assert.Equal(expectedWeight, margin.FontDescriptor.Weight);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void HtmlPageMarginInheritance_ExplicitRegisteredInheritanceOverridesTheNonInheritedDefault(bool inheritOnPage) {
        string page = inheritOnPage ? "--n:inherit;font-size:var(--n)" : "--n:30px;font-size:20px";
        string margin = inheritOnPage ? "" : "--n:inherit;font-size:var(--n)";
        string html = "<style>@property --n{syntax:'<length>';inherits:false;initial-value:12px}</style>"
            + PageMarginInheritanceSource("font:20px/1 Pinned;--n:30px", page, margin);

        HtmlRenderText text = PageMarginTypographyText(Assert.Single(HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse(html), PageMarginTypographyOptions()).Pages));

        Assert.Equal(30D, text.Font.Size, 6);
    }

    [Fact]
    public void HtmlPageMarginInheritance_RegisteredInheritanceAlsoUsesTheSharedDomAndPseudoOwner() {
        const string html = """
            <style>@property --n{syntax:'<length>';inherits:false;initial-value:12px}
              body,p{margin:0;font:16px/1 Pinned}#parent{--n:30px}
              #child{--n:inherit;font-size:var(--n)}
              #child::before{--n:inherit;font-size:var(--n);content:'Pseudo'}
            </style><p id='parent'><span id='child'>Child</span></p>
            """;
        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html),
            PageMarginTypographyOptions()).Pages);
        HtmlRenderText[] text = EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>().ToArray();

        Assert.Equal(30D, Assert.Single(text, item => item.Text == "Child").Font.Size, 6);
        Assert.Equal(30D, Assert.Single(text, item => item.Text == "Pseudo").Font.Size, 6);
    }
    [Theory]
    [InlineData(HtmlRenderUserAgentStyleMode.Document, "font:initial")]
    [InlineData(HtmlRenderUserAgentStyleMode.Document, "all:initial")]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser, "font:initial")]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser, "all:initial")]
    public void HtmlPageMarginInheritance_OrdinaryFontResetsUseTheSameInitialOwner(
        HtmlRenderUserAgentStyleMode mode, string reset) {
        HtmlRenderOptions options = PageMarginTypographyOptions();
        options.UserAgentStyles = mode;
        options.DefaultFontFamily = "Pinned";
        string html = "<style>body,p{margin:0;font:20px Arial;font-stretch:75%;font-weight:500}"
            + "span{" + reset + "}</style><p><span>Reset</span></p>";
        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options).Pages);
        HtmlRenderText text = Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>(), item => item.Text == "Reset");

        Assert.Equal("Pinned", text.Font.FamilyName);
        Assert.Equal(100D, text.FontDescriptor.StretchPercent, 6);
        Assert.Equal(400, text.FontDescriptor.Weight);
    }

}
