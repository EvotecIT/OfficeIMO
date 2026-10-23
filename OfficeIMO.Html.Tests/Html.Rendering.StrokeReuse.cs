using AngleSharp.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlStrokeReuseTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EquivalentStrokeValuesKeepIndependentCascadeWinnersAcrossElements(bool layered) {
        var document = HtmlConversionDocument.Parse("<div id='first'></div><div id='second'></div>")
            .CreateDocumentForRendering();
        IElement firstElement = document.QuerySelector("#first")!;
        IElement secondElement = document.QuerySelector("#second")!;
        var properties = new Dictionary<string, string> {
            ["border"] = "1px solid black", ["border-left-width"] = "2px"
        };
        HtmlCssCascadePriority Low() => new(false, false, false,
            layered ? new CascadeLayerOrder(new[] { 1 }) : null, 0, 0, 0, 1, 1);
        HtmlCssCascadePriority High() => new(false, false, false,
            layered ? new CascadeLayerOrder(new[] { 2 }) : null, 0, 0, 0, layered ? 1 : 2, 1);
        var first = new HtmlComputedStyle(properties, cascadePriorities: new Dictionary<string, HtmlCssCascadePriority> {
            ["border"] = Low(), ["border-left-width"] = High()
        });
        var second = new HtmlComputedStyle(properties, cascadePriorities: new Dictionary<string, HtmlCssCascadePriority> {
            ["border"] = High(), ["border-left-width"] = Low()
        });
        var styles = new HtmlComputedStyleSet(new Dictionary<IElement, HtmlComputedStyle> {
            [firstElement] = first, [secondElement] = second
        }, new Dictionary<IElement, HtmlPseudoElementStylePair>());
        var resolver = new HtmlRenderStyleResolver(styles, new HtmlRenderOptions(), new HtmlDiagnosticReport(), _ => double.NaN);

        Assert.Equal(2D, resolver.Resolve(firstElement, 200D).Borders.Left.Width);
        Assert.Equal(1D, resolver.Resolve(secondElement, 200D).Borders.Left.Width);
        Assert.Equal(2D, resolver.Resolve(firstElement, 200D).Borders.Left.Width);
    }

    [Fact]
    public void RepeatedStyleResolutionRetainsCurrentFontViewportColorAndCharacterMetrics() {
        IElement element = HtmlConversionDocument.Parse("<div></div>").CreateDocumentForRendering().QuerySelector("div")!;
        var computed = new HtmlComputedStyle(new Dictionary<string, string> {
            ["font-size"] = "inherit", ["color"] = "inherit",
            ["border"] = "1em solid currentcolor", ["outline"] = "1vw dashed currentcolor",
            ["outline-offset"] = "1ch"
        }, inheritedProperties: new[] { "font-size", "color" });
        var styles = new HtmlComputedStyleSet(new Dictionary<IElement, HtmlComputedStyle> { [element] = computed },
            new Dictionary<IElement, HtmlPseudoElementStylePair>());
        double characterFactor = 0.5D;
        var resolver = new HtmlRenderStyleResolver(styles, new HtmlRenderOptions(), new HtmlDiagnosticReport(),
            style => style.Font.Size * characterFactor);
        var parent = new HtmlRenderBoxStyle {
            Font = new OfficeFontInfo("Arial", 20D), Color = OfficeColor.FromRgb(255, 0, 0)
        };
        resolver.SetViewport(800D, 600D);
        HtmlRenderBoxStyle first = resolver.Resolve(element, 200D, parent);
        HtmlRenderBorderEdges original = first.Borders;
        first.Borders = HtmlRenderBorderEdges.Uniform(99D, "solid", OfficeColor.Black);
        Assert.Equal(original, resolver.Resolve(element, 200D, parent).Borders);

        parent.Font = parent.Font.WithSize(40D);
        parent.Color = OfficeColor.FromRgb(0, 0, 255);
        characterFactor = 0.75D;
        resolver.SetViewport(1000D, 700D);
        HtmlRenderBoxStyle changed = resolver.Resolve(element, 300D, parent);
        Assert.NotEqual(original.Top.Width, changed.Borders.Top.Width);
        Assert.Equal(parent.Color, changed.Borders.Top.Color);
        Assert.Equal(10D, changed.OutlineWidth, 5);
        Assert.Equal(changed.Font.Size * characterFactor, changed.OutlineOffset, 5);

        Assert.True(HtmlCssBoxStrokeParser.TryParseBorder(computed, 300D, changed.Font.Size,
            resolver.RootFontSize, 1000D, 700D, changed.Color, out HtmlRenderBorderEdges expected,
            out string detail, changed.CharacterAdvance), detail);
        Assert.Equal(expected, changed.Borders);
    }

    [Fact]
    public void UnsupportedStrokeValuesRetainTheirDetailsAcrossRepeatedResolution() {
        IElement element = HtmlConversionDocument.Parse("<div></div>").CreateDocumentForRendering().QuerySelector("div")!;
        var computed = new HtmlComputedStyle(new Dictionary<string, string> {
            ["border"] = "unsupported(2) solid red", ["outline"] = "unsupported(3) solid blue"
        });
        var styles = new HtmlComputedStyleSet(new Dictionary<IElement, HtmlComputedStyle> { [element] = computed },
            new Dictionary<IElement, HtmlPseudoElementStylePair>());
        var resolver = new HtmlRenderStyleResolver(styles, new HtmlRenderOptions(), new HtmlDiagnosticReport(), _ => double.NaN);
        HtmlRenderBoxStyle first = resolver.Resolve(element, 200D);
        HtmlRenderBoxStyle again = resolver.Resolve(element, 200D);
        Assert.NotEmpty(first.UnsupportedBorderPaint);
        Assert.NotEmpty(first.UnsupportedOutlinePaint);
        Assert.Equal(first.UnsupportedBorderPaint, again.UnsupportedBorderPaint);
        Assert.Equal(first.UnsupportedOutlinePaint, again.UnsupportedOutlinePaint);
        Assert.Equal(first.Borders, again.Borders);
    }
}
