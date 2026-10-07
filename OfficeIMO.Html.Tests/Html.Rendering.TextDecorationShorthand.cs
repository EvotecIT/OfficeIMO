using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("underline 0.05em dashed rgb(88,88,88)", "0.05em", "dashed")]
    [InlineData("red 5% dotted underline overline line-through", "5%", "dotted")]
    [InlineData("from-font underline double red", "from-font", "double")]
    [InlineData("underline calc(0.05em + 1px) dashed red", "calc(0.05em + 1px)", "dashed")]
    [InlineData("underline -1px dashed red", "-1px", "dashed")]
    public void TextDecorationThicknessDoesNotDiscardOtherShorthandComponents(string declaration, string thickness, string style) {
        foreach (bool inline in new[] { true, false }) {
            string html = inline
                ? "<p id='target' style='text-decoration:" + declaration + "'>Label</p>"
                : "<style>#target {text-decoration:" + declaration + "}</style><p id='target'>Label</p>";
            HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
            HtmlComputedStyle computed = HtmlComputedStyleEngine.Compute(source)[source.Document.QuerySelector("#target")!];
            Assert.Contains("underline", computed.GetValue("text-decoration-line"));
            Assert.Equal(style, computed.GetValue("text-decoration-style"));
            Assert.Equal(thickness, computed.GetValue("text-decoration-thickness"));
        }
    }

    [Fact]
    public void TextDecorationThicknessFollowsShorthandOrderImportantVariablesAndReset() {
        const string html = """
            <style>
              .target { text-decoration:underline 4px dotted red !important; }
              #variable { --decoration:underline 5% dashed red; text-decoration:var(--decoration); }
            </style>
            <p id='reset' style='text-decoration-thickness:4px;text-decoration:underline'>Reset</p>
            <p id='later' style='text-decoration:underline 4px dotted red;text-decoration-thickness:2px'>Later</p>
            <p id='important' class='target' style='text-decoration-thickness:2px'>Important</p>
            <p id='variable'>Variable</p>
            <p id='invalid' style='text-decoration:underline 4px dotted red;text-decoration:var(--missing)'>Invalid</p>
            <p id='wide' style='text-decoration-thickness:4px;text-decoration:initial'>Wide</p>
            """;
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        var styles = HtmlComputedStyleEngine.Compute(source);
        string Value(string id) => styles[source.Document.QuerySelector("#" + id)!].GetValue("text-decoration-thickness");
        Assert.Equal("auto", Value("reset"));
        Assert.Equal("2px", Value("later"));
        Assert.Equal("4px", Value("important"));
        Assert.Equal("5%", Value("variable"));
        Assert.Equal(string.Empty, Value("invalid"));
        Assert.Equal(string.Empty, Value("wide"));
    }

    [Theory]
    [InlineData("underline 2 dashed red")]
    [InlineData("underline 1px 2px dashed red")]
    [InlineData("underline auto from-font red")]
    [InlineData("underline none 1px")]
    [InlineData("underline underline 1px")]
    public void InvalidTextDecorationShorthandsPreserveEarlierDeclaration(string invalid) {
        foreach (bool inline in new[] { true, false }) {
            string declarations = "text-decoration:underline 3px dotted blue;text-decoration:" + invalid;
            string html = inline ? "<p id='target' style='" + declarations + "'>Label</p>"
                : "<style>#target {" + declarations + "}</style><p id='target'>Label</p>";
            HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
            HtmlComputedStyle computed = HtmlComputedStyleEngine.Compute(source)[source.Document.QuerySelector("#target")!];
            Assert.Equal("dotted", computed.GetValue("text-decoration-style"));
            Assert.Equal("3px", computed.GetValue("text-decoration-thickness"));
        }
    }

    [Fact]
    public void ThicknessContainingLinkPaintPreservesAuthoredStyleColorAndReportsApproximation() {
        const string html = """
            <a href='https://example.test' style='text-decoration:underline 0.05em dashed rgb(88,88,88)'><span>Decorated label</span></a>
            <span style='text-decoration:none 2px red'>No decoration</span>
            <a href='https://example.test/auto' style='text-decoration:underline dashed red'>Automatic label</a>
            """;
        HtmlRenderDocument scene = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 800 });
        HtmlRenderText label = Assert.Single(scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).OfType<HtmlRenderText>(), t => t.Text == "Decorated label");
        Assert.Equal(OfficeTextDecorationStyle.Dashed, label.UnderlineStyle);
        Assert.Equal(OfficeColor.FromRgb(88,88,88), label.DecorationColor);
        HtmlDiagnostic diagnostic = Assert.Single(scene.Diagnostics, d => d.Code == "HtmlRenderTextDecorationThicknessApproximated");
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.True(scene.HasLoss);
    }
}
