using AngleSharp.Html.Parser;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlFontShorthandTests {
    [Theory]
    [InlineData("font:12px Arial", "", "")]
    [InlineData("--type:12px Arial;font:var(--type)", "", "")]
    [InlineData("font:12px Arial;font-kerning:none", "none", "")]
    [InlineData("font:inherit", "normal", "tabular-nums")]
    [InlineData("font:unset", "normal", "tabular-nums")]
    public void FontShorthandResetsSupportedFeaturesAndPreservesIndependentPalette(string declaration, string kerning, string numeric) {
        foreach (bool inline in new[] { true, false }) {
            string declarations = "font-kerning:none;font-variant-numeric:oldstyle-nums;font-feature-settings:\"liga\" 0;font-palette:dark;" + declaration;
            var document = new HtmlParser().ParseDocument((inline ? "" : "<style>#target{" + declarations + "}</style>") +
                "<div style='font-kerning:normal;font-variant-numeric:tabular-nums'><p id='target'" +
                (inline ? " style='" + declarations + "'" : "") + ">Text</p></div>");
            var computed = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];
            Assert.Equal(kerning, computed.GetValue("font-kerning"));
            Assert.Equal(numeric, computed.GetValue("font-variant-numeric"));
            Assert.Equal("", computed.GetValue("font-feature-settings"));
            Assert.Equal("dark", computed.GetValue("font-palette"));
        }
    }

    [Theory]
    [InlineData("font:italic bold 6px/8px Arial", "6px", "8px", "italic", "bold")]
    [InlineData("font:6px/8px Arial;font-size:10px", "10px", "8px", "normal", "normal")]
    [InlineData("font-size:10px;font:6px/8px Arial", "6px", "8px", "normal", "normal")]
    [InlineData("font-size:10px!important;font:6px/8px Arial", "10px", "8px", "normal", "normal")]
    [InlineData("font:6px/8px Arial!important;font-size:10px", "6px", "8px", "normal", "normal")]
    [InlineData("--type:6px/8px Arial;font:var(--type);font-size:10px", "10px", "8px", "normal", "normal")]
    [InlineData("--type:6px/8px Arial;font-size:10px;font:var(--type)", "6px", "8px", "normal", "normal")]
    [InlineData("font-size:10px;font:6px/8px Arial;font-size:12px", "12px", "8px", "normal", "normal")]
    [InlineData("font-size:10px!important;font:6px/8px Arial;font-size:12px", "10px", "8px", "normal", "normal")]
    [InlineData("font:6px/8px Arial;font-size:10px;font:12px/8px Arial", "12px", "8px", "normal", "normal")]
    public void InlineFontShorthandKeepsItsLonghandCascade(string declarations, string size, string height, string style, string weight) {
        foreach (bool inline in new[] { true, false }) {
            string html = (inline ? "" : "<style>#target{" + declarations + "}</style>") +
                "<div style='font-style:italic;font-weight:bold'><p id='target'" + (inline ? " style='" + declarations + "'" : "") + ">Text</p></div>";
            var document = new HtmlParser().ParseDocument(html);
            var computed = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];
            Assert.Equal(size, computed.GetValue("font-size"));
            Assert.Equal(height, computed.GetValue("line-height"));
            Assert.Equal(style, computed.GetValue("font-style"));
            Assert.Equal(weight, computed.GetValue("font-weight"));
            Assert.Equal("Arial", computed.GetValue("font-family"));
        }
    }

    [Theory]
    [InlineData("font:inherit", "18px", "italic")]
    [InlineData("font:initial", "", "")]
    [InlineData("font:bad-value", "18px", "italic")]
    public void InlineFontShorthandRespectsInheritedAndInvalidValues(string declaration, string size, string style) {
        var document = new HtmlParser().ParseDocument("<div style='font-size:18px;font-style:italic'><span id='target' style='" + declaration + "'>Text</span></div>");
        var computed = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];
        Assert.Equal(size, computed.GetValue("font-size"));
        Assert.Equal(style, computed.GetValue("font-style"));
    }

    [Theory]
    [InlineData("font-weight:bold;font-weight:banana", "18px", "bold")]
    [InlineData("font-weight:banana!important;font-weight:bold", "18px", "bold")]
    [InlineData("font:20px Arial;font:banana", "20px", "normal")]
    [InlineData("font:banana!important;font:20px Arial", "20px", "normal")]
    [InlineData("font-size:30px;font:var(--missing)", "18px", "normal")]
    [InlineData("font:var(--missing);font-size:30px", "30px", "normal")]
    [InlineData("font-size:30px!important;font:var(--missing)", "30px", "normal")]
    [InlineData("font-size:30px;font:var(--missing)!important", "18px", "normal")]
    [InlineData("--type:banana;font-size:30px;font:var(--type)", "18px", "normal")]
    [InlineData("font-size:30px;font:var(--missing, 20px Arial)", "20px", "normal")]
    public void InvalidFontDeclarationsRespectParseAndComputedValueBoundaries(string declarations, string size, string weight) {
        foreach (bool inline in new[] { true, false }) {
            string html = (inline ? "" : "<style>#target{" + declarations + "}</style>") +
                "<div style='font-size:18px;font-weight:normal'><p id='target'" + (inline ? " style='" + declarations + "'" : "") + ">Text</p></div>";
            var document = new HtmlParser().ParseDocument(html);
            var computed = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];
            Assert.Equal(size, computed.GetValue("font-size"));
            Assert.Equal(weight, computed.GetValue("font-weight"));
        }
    }

    [Fact]
    public void FontLonghandRevertLayerRetainsTheEarlierLayersFontVariable() {
        var document = new HtmlParser().ParseDocument("<style>@layer base, override;@layer base{#target{--type:20px Arial;font:var(--type)}}@layer override{#target{font:30px Arial;font-size:revert-layer}}</style><p id='target'>Text</p>");
        var computed = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];
        Assert.Equal("20px", computed.GetValue("font-size"));
    }

    [Theory]
    [InlineData("font-size:20px;font-size:banana", "20px", "24px")]
    [InlineData("font-size:banana!important;font-size:20px", "20px", "24px")]
    [InlineData("line-height:24px;line-height:banana", "18px", "24px")]
    [InlineData("line-height:banana!important;line-height:24px", "18px", "24px")]
    public void InvalidFontLengthsDoNotReplaceValidDeclarations(string declarations, string size, string height) {
        foreach (bool inline in new[] { true, false }) {
            var document = new HtmlParser().ParseDocument((inline ? "" : "<style>#target{" + declarations + "}</style>") +
                "<div style='font-size:18px;line-height:24px'><p id='target'" + (inline ? " style='" + declarations + "'" : "") + ">Text</p></div>");
            var computed = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];
            Assert.Equal(size, computed.GetValue("font-size"));
            Assert.Equal(height, computed.GetValue("line-height"));
        }
    }
}
