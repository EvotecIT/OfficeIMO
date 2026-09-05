using AngleSharp.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlCustomProperties_RegisteredNonInheritedValuesUseTheirInitialValueInChildren() {
        const string html = """
            <style>
              @property --tone { syntax: "<color>"; inherits: false; initial-value: red; }
              .parent { --tone: blue; }
              .child { color: var(--tone); }
            </style>
            <div class="parent"><span class="child">Child</span></div>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("blue", styles[document.QuerySelector(".parent")!].GetValue("--tone"));
        Assert.Equal("red", styles[document.QuerySelector(".child")!].GetValue("--tone"));
        Assert.Equal("red", styles[document.QuerySelector(".child")!].GetValue("color"));
    }

    [Fact]
    public void HtmlCustomProperties_RegisteredInheritedValuesHonorUnsetAndInvalidComputedFallbacks() {
        const string html = """
            <style>
              @property --space { syntax: "<length>"; inherits: true; initial-value: 3px; }
              .parent { --space: 11px; }
              .unset { --space: unset; padding-left: var(--space); }
              .invalid { --space: tomato; padding-left: var(--space); }
            </style>
            <div class="parent"><i class="unset"></i><b class="invalid"></b></div>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("11px", styles[document.QuerySelector(".unset")!].GetValue("--space"));
        Assert.Equal("11px", styles[document.QuerySelector(".unset")!].GetValue("padding-left"));
        Assert.Equal("11px", styles[document.QuerySelector(".invalid")!].GetValue("--space"));
        Assert.Equal("11px", styles[document.QuerySelector(".invalid")!].GetValue("padding-left"));
    }

    [Fact]
    public void HtmlCustomProperties_IgnoreInvalidRegistrationsAndAcceptTypedLists() {
        const string html = """
            <style>
              @property --broken { syntax: "<length>"; inherits: false; initial-value: red; }
              @property --stops { syntax: "<percentage>#"; inherits: false; initial-value: 10%, 90%; }
              .parent { --broken: 8px; --stops: 20%, 80%; }
              .child { width: var(--broken); }
            </style>
            <div class="parent"><span class="child"></span></div>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("8px", styles[document.QuerySelector(".child")!].GetValue("--broken"));
        Assert.Equal("8px", styles[document.QuerySelector(".child")!].GetValue("width"));
        Assert.Equal("10%, 90%", styles[document.QuerySelector(".child")!].GetValue("--stops"));
    }

    [Fact]
    public void HtmlCustomProperties_RegisteredValuesFlowIntoPseudoElementsAndRendering() {
        const string html = """
            <style>
              @property --accent { syntax: "<color>"; inherits: true; initial-value: black; }
              p { --accent: #123456; }
              p::before { content: "Registered"; color: var(--accent); }
            </style>
            <p>Body</p>
            """;

        HtmlRenderText generated = Assert.Single(
            HtmlRenderTestDriver.Render(html).Pages[0].Visuals.OfType<HtmlRenderText>(),
            item => item.Text == "Registered");
        Assert.Equal(OfficeColor.FromRgb(0x12, 0x34, 0x56), generated.Color);
    }
}
