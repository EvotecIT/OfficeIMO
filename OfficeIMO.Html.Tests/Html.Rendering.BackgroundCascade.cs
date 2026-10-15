using AngleSharp.Html.Parser;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlBackgroundCascadeTests {
    [Theory]
    [InlineData("background:lime", "background:red", "red")]
    [InlineData("background-color:lime", "background:red", "red")]
    [InlineData("background:lime!important", "background:red", "lime")]
    [InlineData("background:lime", "background:red!important", "red")]
    [InlineData("background:lime!important", "background:red!important", "red")]
    public void InlineBackgroundExpansionRetainsDeclarationPriority(string stylesheet, string inline, string color) {
        HtmlComputedStyle computed = Compute("<style>#target{" + stylesheet + "}</style><div id='target' style='" + inline + "'></div>");
        Assert.Equal(OfficeColor.ParseCss(color), OfficeColor.ParseCss(computed.GetValue("background-color")));
        Assert.Equal("none", computed.GetValue("background-image"));
    }

    [Theory]
    [InlineData("background:red;background-color:blue", "blue")]
    [InlineData("background-color:blue;background:red", "red")]
    [InlineData("background:red!important;background-color:blue", "red")]
    [InlineData("background-color:blue!important;background:red", "blue")]
    [InlineData("background:red;background-color:blue;background:lime", "lime")]
    [InlineData("background:red;background-color:blue;background:lime!important;background-color:black", "lime")]
    [InlineData("--paint:red;background:var(--paint);background-color:blue", "blue")]
    [InlineData("--paint:red;background-color:blue;background:var(--paint)", "red")]
    [InlineData("background:red;background:var(--missing)", "transparent")]
    [InlineData("background:red;background-color:var(--missing)", "transparent")]
    [InlineData("background:red;background-color:initial", "transparent")]
    [InlineData("background:red;background-color:unset", "transparent")]
    public void BackgroundComponentsKeepAuthoredOrderAndComputedResets(string declaration, string color) {
        foreach (bool inline in new[] { true, false }) {
            HtmlComputedStyle computed = Compute(ForDeclaration(declaration, inline));
            Assert.Equal(OfficeColor.ParseCss(color), OfficeColor.ParseCss(computed.GetValue("background-color")));
        }
    }

    [Theory]
    [InlineData("background:red;background-color:initial", "transparent", "none", "0% 0%", "auto", "repeat", "scroll", "padding-box", "border-box")]
    [InlineData("background:initial", "transparent", "none", "0% 0%", "auto", "repeat", "scroll", "padding-box", "border-box")]
    [InlineData("background:inherit", "blue", "none", "right bottom", "20px 10px", "no-repeat", "local", "content-box", "padding-box")]
    [InlineData("background-size:20px 10px!important;background:red", "red", "none", "0% 0%", "20px 10px", "repeat", "scroll", "padding-box", "border-box")]
    public void BackgroundShorthandResetsEveryOmittedComponentAndInheritsComputedLonghands(
        string declaration, string color, string image, string position, string size, string repeat, string attachment, string origin, string clip) {
        const string earlier = "background:blue;background-position:right bottom;background-size:20px 10px;background-repeat:no-repeat;background-attachment:local;background-origin:content-box;background-clip:padding-box;";
        foreach (bool inline in new[] { true, false }) {
            HtmlComputedStyle computed = Compute("<div style='" + earlier + "'>" + ForDeclaration(earlier + declaration, inline) + "</div>");
            Assert.Equal(OfficeColor.ParseCss(color), OfficeColor.ParseCss(computed.GetValue("background-color")));
            Assert.Equal(image, computed.GetValue("background-image"));
            Assert.Equal(position, computed.GetValue("background-position"));
            Assert.Equal(size, computed.GetValue("background-size"));
            Assert.Equal(repeat, computed.GetValue("background-repeat"));
            Assert.Equal(attachment, computed.GetValue("background-attachment"));
            Assert.Equal(origin, computed.GetValue("background-origin"));
            Assert.Equal(clip, computed.GetValue("background-clip"));
        }
    }

    [Fact]
    public void BackgroundExpansionPreservesEmptyLayersAndSingleVisualBoxes() {
        const string declaration = "background:linear-gradient(red,blue) left top / 10px 12px no-repeat content-box, lime padding-box";
        foreach (bool inline in new[] { true, false }) {
            HtmlComputedStyle computed = Compute(ForDeclaration(declaration, inline));
            Assert.EndsWith(", none", computed.GetValue("background-image"), StringComparison.Ordinal);
            Assert.Equal("10px 12px, auto", computed.GetValue("background-size"));
            Assert.Equal("no-repeat, repeat", computed.GetValue("background-repeat"));
            Assert.Equal("content-box, padding-box", computed.GetValue("background-origin"));
            Assert.Equal("content-box, padding-box", computed.GetValue("background-clip"));
            Assert.Equal(OfficeColor.ParseCss("lime"), OfficeColor.ParseCss(computed.GetValue("background-color")));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BackgroundExpansionRetainsOwnedGradientAndColorSyntax(bool variable) {
        const string gradient = "repeating-linear-gradient(to right,rgba(255 0 0 / 50%) 0 10px,color-mix(in srgb,blue 75%,white) 10px 20px)";
        const string color = "color-mix(in srgb, red 25%, blue)";
        string value = gradient + " center / 20px 10px no-repeat " + color;
        string declaration = variable ? "--paint:" + value + ";background:var(--paint)" : "background:" + value;
        foreach (bool inline in new[] { true, false }) {
            HtmlComputedStyle computed = Compute(ForDeclaration(declaration, inline));
            Assert.Equal(gradient, computed.GetValue("background-image"));
            Assert.Equal("20px 10px", computed.GetValue("background-size"));
            Assert.Equal(OfficeColor.FromRgb(64, 0, 191), OfficeColor.ParseCss(computed.GetValue("background-color")));
        }
    }

    [Theory]
    [InlineData("image-set(url(\"https://assets.example.test/x.png\") 1x)")]
    [InlineData("-webkit-image-set(url(\"https://assets.example.test/x.png\") 1x)")]
    [InlineData("image-set(url(\"https://assets.example.test/x.png\") type(\"\"))")]
    [InlineData("image-set(\"\")")]
    [InlineData("paint(example)")]
    public void BackgroundExpansionRetainsImageFunctionsAndPendingResourcesThroughVariables(string image) {
        foreach (bool inline in new[] { true, false }) {
            foreach (bool variable in new[] { true, false }) {
                string declaration = "width:20px;height:12px;"
                    + (variable ? "--paint:" + image + " red;background:var(--paint)" : "background:" + image + " red");
                string html = "<style>#target{background:lime}</style>" + ForDeclaration(declaration, inline);
                HtmlComputedStyle computed = Compute(html);
                Assert.Equal(image, computed.GetValue("background-image"));
                Assert.Equal(OfficeColor.Red, OfficeColor.ParseCss(computed.GetValue("background-color")));
                HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
                Assert.Equal(OfficeColor.Red, Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
                    shape => shape.Source == "div#target").Shape.FillColor);
                if (image.IndexOf("url(", StringComparison.Ordinal) >= 0) {
                    HtmlDiagnostic pending = Assert.Single(rendered.Diagnostics,
                        diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
                    Assert.Equal("https://assets.example.test/x.png", pending.Detail);
                    Assert.Equal("div#target:background-image", pending.Source);
                }
            }
        }
    }

    [Theory]
    [InlineData("initial,red")]
    [InlineData("inherit,red")]
    [InlineData("unset,red")]
    [InlineData("red,")]
    [InlineData("red nonsense")]
    public void InvalidBackgroundLiteralsAreIgnoredButInvalidSubstitutionResetsComponents(string value) {
        foreach (bool inline in new[] { true, false }) {
            foreach (bool variable in new[] { true, false }) {
                string declaration = variable ? "--paint:" + value + ";background:var(--paint)" : "background:" + value;
                HtmlComputedStyle computed = Compute("<style>#target{background:blue}</style>" + ForDeclaration(declaration, inline));
                Assert.Equal(variable ? OfficeColor.Transparent : OfficeColor.Blue,
                    OfficeColor.ParseCss(computed.GetValue("background-color")));
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InvalidBackgroundWithoutAnEarlierDeclarationCannotPaintAParsedLayer(bool inline) {
        string html = ForDeclaration("width:20px;height:12px;background:initial,red", inline);
        Assert.Equal("", Compute(html).GetValue("background-color"));
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#target");
    }

    [Theory]
    [InlineData("background-color", "red", "transparent")]
    [InlineData("background-image", "linear-gradient(red,blue)", "none")]
    [InlineData("background-position", "right bottom", "0% 0%")]
    [InlineData("background-size", "20px 10px", "auto")]
    [InlineData("background-repeat", "no-repeat", "repeat")]
    [InlineData("background-attachment", "fixed", "scroll")]
    [InlineData("background-origin", "content-box", "padding-box")]
    [InlineData("background-clip", "content-box", "border-box")]
    public void InvalidBackgroundLonghandsRespectLiteralAndComputedValueBoundaries(string property, string valid, string initial) {
        foreach (bool inline in new[] { true, false }) {
            foreach (bool variable in new[] { true, false }) {
                string declarations = property + ":" + valid + ";"
                    + (variable ? "--invalid:bogus;" + property + ":var(--invalid)" : property + ":bogus");
                string actual = Compute(ForDeclaration(declarations, inline)).GetValue(property);
                string expected = variable ? initial : valid;
                if (property == "background-color") Assert.Equal(OfficeColor.ParseCss(expected), OfficeColor.ParseCss(actual));
                else Assert.Equal(expected, actual);
            }
        }
    }

    [Theory]
    [InlineData("linear-gradient(bogus)")]
    [InlineData("radial-gradient(bogus)")]
    [InlineData("conic-gradient(bogus)")]
    [InlineData("image-set()")]
    [InlineData("-webkit-image-set()")]
    [InlineData("paint()")]
    [InlineData("image-set(linear-gradient(bogus) 1x)")]
    [InlineData("image-set(image-set(\"x.png\" 1x) 1x)")]
    [InlineData("image-set(-webkit-image-set(\"x.png\" 1x) 1x)")]
    [InlineData("-webkit-image-set(image-set(\"x.png\" 1x) 1x)")]
    [InlineData("-webkit-image-set(-webkit-image-set(\"x.png\" 1x) 1x)")]
    public void OpaqueBackgroundFunctionsCannotHideInvalidImageSyntax(string image) {
        foreach (bool inline in new[] { true, false }) {
            foreach (bool variable in new[] { true, false }) {
                string value = image + " blue";
                string declarations = "background:red;"
                    + (variable ? "--invalid:" + value + ";background:var(--invalid)" : "background:" + value);
                HtmlComputedStyle computed = Compute(ForDeclaration(declarations, inline));
                Assert.Equal(variable ? OfficeColor.Transparent : OfficeColor.Red,
                    OfficeColor.ParseCss(computed.GetValue("background-color")));
                Assert.Equal("none", computed.GetValue("background-image"));
            }
        }
    }

    [Theory]
    [InlineData("background-color", "color-mix(in srgb, red 25%, blue)")]
    [InlineData("background-position", "calc(50% - 2px) 0")]
    [InlineData("background-size", "min(10px,20px) 20cqh")]
    public void BackgroundLonghandAdmissionRetainsOwnedColorAndMathValues(string property, string value) {
        foreach (bool inline in new[] { true, false }) {
            foreach (bool variable in new[] { true, false }) {
                string declaration = variable ? "--value:" + value + ";" + property + ":var(--value)" : property + ":" + value;
                Assert.Equal(value, Compute(ForDeclaration(declaration, inline)).GetValue(property));
            }
        }
    }

    [Fact]
    public void BackgroundShorthandPreservesOwnedPositionAndSizeMathThroughProviderExpansion() {
        const string value = "linear-gradient(red,blue) calc(50% - 2px) 0 / min(10px,20px) 20cqh no-repeat color-mix(in srgb, red 25%, blue)";
        foreach (bool inline in new[] { true, false }) {
            foreach (bool variable in new[] { true, false }) {
                string declaration = variable ? "--paint:" + value + ";background:var(--paint)" : "background:" + value;
                HtmlComputedStyle computed = Compute(ForDeclaration(declaration, inline));
                Assert.Contains("calc(50% - 2px)", computed.GetValue("background-position"), StringComparison.Ordinal);
                Assert.Equal("min(10px,20px) 20cqh", computed.GetValue("background-size"));
                Assert.Equal(OfficeColor.FromRgb(64, 0, 191), OfficeColor.ParseCss(computed.GetValue("background-color")));
            }
        }
    }

    [Fact]
    public void BackgroundSyntaxAdmissionDoesNotConsumeTheGradientRenderStopLimit() {
        const string html = "<div style='width:40px;height:20px;background:linear-gradient(red,lime,blue) blue'></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { MaxGradientStops = 2 });
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GradientStopLimitExceeded);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.FillColor == OfficeColor.Blue);
    }

    [Theory]
    [InlineData("background:revert-layer", "blue", "20px 10px")]
    [InlineData("background:red;background-color:revert-layer", "blue", "auto")]
    public void BackgroundLayerRevertKeepsEachLonghandsEarlierCandidate(string declaration, string color, string size) {
        HtmlComputedStyle computed = Compute("<style>@layer base,theme;@layer base{#target{background:blue;background-size:20px 10px}}"
            + "@layer theme{#target{" + declaration + "}}</style><div id='target'></div>");
        Assert.Equal(OfficeColor.ParseCss(color), OfficeColor.ParseCss(computed.GetValue("background-color")));
        Assert.Equal(size, computed.GetValue("background-size"));
    }

    [Theory]
    [InlineData("background:red", "red")]
    [InlineData("background:red;background-color:blue", "blue")]
    [InlineData("background:red;background-color:initial", "transparent")]
    [InlineData("background:red;background-color:var(--missing)", "transparent")]
    public void BackgroundCascadeControlsActualTargetPaint(string inline, string color) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render("<style>.target{background:lime}</style><div id='target' class='target' style='width:20px;height:12px;" + inline + "'></div>");
        IEnumerable<HtmlRenderShape> backgrounds = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>().Where(shape => shape.Source == "div#target");
        if (color == "transparent") Assert.Empty(backgrounds);
        else Assert.Equal(OfficeColor.ParseCss(color), Assert.Single(backgrounds).Shape.FillColor);
    }

    private static string ForDeclaration(string declaration, bool inline) =>
        (inline ? "" : "<style>#target{" + declaration + "}</style>")
        + "<div id='target'" + (inline ? " style='" + declaration + "'" : "") + "></div>";

    private static HtmlComputedStyle Compute(string html) {
        var document = new HtmlParser().ParseDocument(html);
        return HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];
    }
}
