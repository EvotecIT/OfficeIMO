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
    [Theory]
    [InlineData("0.0")]
    [InlineData("+0")]
    [InlineData("-0")]
    public void EquivalentZeroThicknessSpellingsPreserveDecoration(string zero) {
        foreach (bool inline in new[] { true, false }) {
            string css = "text-decoration:underline " + zero + " dashed red";
            string html = inline ? "<p id='target' style='" + css + "'>Label</p>"
                : "<style>#target {" + css + "}</style><p id='target'>Label</p>";
            HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
            HtmlComputedStyle computed = HtmlComputedStyleEngine.Compute(source)[source.Document.QuerySelector("#target")!];
            Assert.Equal("underline", computed.GetValue("text-decoration-line"));
            Assert.Equal("dashed", computed.GetValue("text-decoration-style"));
            Assert.Equal(zero, computed.GetValue("text-decoration-thickness"));
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void InvalidDecorationLineListsCannotReplaceValidStrikeThrough(bool inline) {
        foreach (string invalid in new[] { "none underline", "underline underline", "line-through line-through" }) {
            string css = "text-decoration-line:line-through;text-decoration-line:" + invalid;
            string html = inline ? "<p id='target' style='" + css + "'>Label</p>"
                : "<style>#target {" + css + "}</style><p id='target'>Label</p>";
            HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
            HtmlComputedStyle computed = HtmlComputedStyleEngine.Compute(source)[source.Document.QuerySelector("#target")!];
            Assert.Equal("line-through", computed.GetValue("text-decoration-line"));
            HtmlRenderDocument scene = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());
            HtmlRenderText label = Assert.Single(scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).OfType<HtmlRenderText>(), t => t.Text == "Label");
            Assert.Equal(OfficeTextDecorationStyle.Single, label.StrikethroughStyle);
            Assert.Equal(OfficeTextDecorationStyle.None, label.UnderlineStyle);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void InheritedShorthandsKeepParentsIndependentlyComputedLonghands(bool inline) {
        const string parentCss = "text-decoration:underline 2px dashed red;text-decoration-thickness:4px;text-decoration-color:blue;text-decoration-style:dotted;margin:2px;margin-top:7px;border:1px solid red;border-top-width:3px";
        const string childCss = "text-decoration:inherit;margin:inherit;border:inherit";
        string html = inline ? "<div id='parent' style='" + parentCss + "'><span id='child' style='" + childCss + "'>Label</span></div>"
            : "<style>#parent {" + parentCss + "} #child {" + childCss + "}</style><div id='parent'><span id='child'>Label</span></div>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        var styles = HtmlComputedStyleEngine.Compute(source);
        HtmlComputedStyle parent = styles[source.Document.QuerySelector("#parent")!];
        HtmlComputedStyle child = styles[source.Document.QuerySelector("#child")!];
        foreach (string property in new[] { "text-decoration-line", "text-decoration-thickness", "text-decoration-color", "text-decoration-style", "margin-top", "border-top-width" }) {
            Assert.Equal(parent.GetValue(property), child.GetValue(property));
            Assert.True(child.IsInheritedValue(property));
        }
    }

    [Theory]
    [InlineData("overline", false)]
    [InlineData("overline underline line-through", true)]
    public void OverlineOmissionIsReportedSeparatelyFromPaintedThickness(string lines, bool hasPaintedLines) {
        string html = "<span style='text-decoration:" + lines + " 2px dashed red'>Label</span>";
        HtmlRenderDocument scene = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());
        HtmlDiagnostic omission = Assert.Single(scene.Diagnostics, d => d.Code == "HtmlRenderTextDecorationLineUnsupported");
        Assert.Equal(OfficeConversionLossKind.Omission, omission.LossKind);
        Assert.Equal(hasPaintedLines ? 1 : 0, scene.Diagnostics.Count(d => d.Code == "HtmlRenderTextDecorationThicknessApproximated"));
        HtmlRenderText label = Assert.Single(scene.Pages.SelectMany(p => EnumerateCorpusVisuals(p.Scene)).OfType<HtmlRenderText>(), t => t.Text == "Label");
        Assert.Equal(hasPaintedLines ? OfficeTextDecorationStyle.Dashed : OfficeTextDecorationStyle.None, label.UnderlineStyle);
        Assert.Equal(hasPaintedLines ? OfficeTextDecorationStyle.Dashed : OfficeTextDecorationStyle.None, label.StrikethroughStyle);
        Assert.True(scene.HasLoss);
    }

}
