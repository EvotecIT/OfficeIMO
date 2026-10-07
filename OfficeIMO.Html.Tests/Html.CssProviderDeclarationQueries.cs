using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlCssProviderDeclarationQueryTests {
    [Theory]
    [InlineData("display:contents;margin:0")]
    [InlineData("display:contents!important;display:block;-webkit-appearance:none")]
    [InlineData("display:contents;display:made-up!important;-webkit-appearance:none")]
    public void ProviderFallbackPreservesSupportedDisplayWithOtherDeclarations(string declarations) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>#target{" + declarations + "}</style><div id='target'>Content</div>");
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("#target")!];

        Assert.Equal("contents", style.GetValue("display"));
    }

    [Fact]
    public void ProviderFallbackPreservesVariableShorthandsNestedFamiliesAndImportance() {
        // The vendor declaration makes this rule use the retained provider path.
        var document = HtmlDocumentEngine.Default.ParseDocument("""
            <style>
              #target { -webkit-appearance:none; --gap:5px; margin:var(--gap) 7px;
                padding:1px 2px 3px 4px!important; border-top:2px solid red;
                font:italic bold 12px/1.5 Arial; }
              #target { padding-left:99px; }
            </style><div id="target">Content</div>
            """);
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];
        Assert.Equal("5px", style.GetValue("margin-top"));
        Assert.Equal("7px", style.GetValue("margin-right"));
        Assert.Equal("5px", style.GetValue("margin-bottom"));
        Assert.Equal("7px", style.GetValue("margin-left"));
        Assert.Equal("1px", style.GetValue("padding-top"));
        Assert.Equal("2px", style.GetValue("padding-right"));
        Assert.Equal("3px", style.GetValue("padding-bottom"));
        Assert.Equal("4px", style.GetValue("padding-left"));
        Assert.Equal("2px", style.GetValue("border-top-width"));
        Assert.True(OfficeIMO.Drawing.OfficeColor.TryParseCss(style.GetValue("border-top-color"), out var borderColor));
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.Red, borderColor);
        Assert.Equal("italic", style.GetValue("font-style"));
        Assert.Equal("bold", style.GetValue("font-weight"));
        Assert.Equal("12px", style.GetValue("font-size"));
        Assert.Equal("1.5", style.GetValue("line-height"));
        Assert.Equal("Arial", style.GetValue("font-family"));
    }

    [Fact]
    public void ProviderFallbackKeepsEachRulesDeclarationsAndObservesLaterStylesheetChanges() {
        var document = new AngleSharp.Html.Parser.HtmlParser().ParseDocument("""
            <style>
              #first { -webkit-appearance:none; margin:3px 7px; }
              #second { -webkit-appearance:none; color:blue; }
            </style><div id="first">First</div><div id="second">Second</div>
            """);
        var first = document.QuerySelector("#first")!;
        var second = document.QuerySelector("#second")!;
        var initial = HtmlComputedStyleEngine.Compute(document);
        Assert.Equal("3px", initial[first].GetValue("margin-top"));
        Assert.Equal("7px", initial[first].GetValue("margin-right"));
        Assert.Empty(initial[second].GetValue("margin-top"));
        Assert.Equal("rgba(0, 0, 255, 1)", initial[second].GetValue("color"));

        document.QuerySelector("style")!.TextContent = "#first{-webkit-appearance:none;padding-left:11px;}";
        var changed = HtmlComputedStyleEngine.Compute(document);
        Assert.Empty(changed[first].GetValue("margin-top"));
        Assert.Equal("11px", changed[first].GetValue("padding-left"));
        Assert.Empty(changed[second].GetValue("padding-left"));
        Assert.Empty(changed[second].GetValue("color"));
    }
}
