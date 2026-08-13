using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRendering_ViewportUnitsAndLengthMathUseTheConfiguredStaticViewport() {
        const string html = """
            <div id="classic" style="width:50vw;height:10vh;background:red"></div>
            <div id="dynamic" style="width:calc(10dvw + 5svw);height:10lvh;background:blue"></div>
            <div id="minimum" style="width:10vmin;height:5vmax;background:lime"></div>
            """;
        var options = new HtmlRenderOptions {
            ViewportWidth = 200D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape classic = FindShape(rendered, "div#classic");
        HtmlRenderShape dynamic = FindShape(rendered, "div#dynamic");
        HtmlRenderShape minimum = FindShape(rendered, "div#minimum");

        Assert.Equal((100D, 10D), (classic.Width, classic.Height));
        Assert.Equal((30D, 10D), (dynamic.Width, dynamic.Height));
        Assert.Equal((10D, 10D), (minimum.Width, minimum.Height));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(width:10dvw)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(height:calc(10svh + 2px))"));
    }

    [Fact]
    public void HtmlRendering_AspectRatioSizesOrdinaryContentAndBorderBoxes() {
        const string html = """
            <div id="content" style="width:80px;aspect-ratio:2/1;background:red"></div>
            <div id="border" style="box-sizing:border-box;width:80px;aspect-ratio:2/1;padding:10px;border:2px solid blue;background:lime"></div>
            """;
        var options = new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape content = FindShape(rendered, "div#content");
        HtmlRenderShape border = FindShape(rendered, "div#border");

        Assert.Equal((80D, 40D), (content.Width, content.Height));
        Assert.Equal((80D, 40D), (border.Width, border.Height));
    }

    [Fact]
    public void HtmlRendering_ContainerQueriesUseTheNearestEligibleOrNamedAncestor() {
        const string html = """
            <style>
              @container (width > 300px) { .item { background:red; } }
              @container card (300px <= width < 400px) { .item { background:blue; } }
              @container theme style(--density: compact) { .item { color:lime; } }
              @container theme style(color:#ff0000) { .equivalent { text-transform:uppercase; } }
            </style>
            <section style="width:360px;container:card / inline-size;--density:compact;container-name:card theme;color:red">
              <div style="width:200px;container-type:inline-size">
                <div id="item" class="item" style="width:50cqw;height:10cqh;font-size:5cqw">ContainerMarker</div>
                <span class="equivalent">Equivalent</span>
              </div>
            </section>
            """;
        var options = new HtmlRenderOptions {
            ViewportWidth = 420D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        HtmlRenderShape item = FindShape(rendered, "div#item");
        HtmlRenderText text = Assert.Single(
            rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(),
            visual => visual.Text == "ContainerMarker");

        Assert.Equal(OfficeColor.Blue, item.Shape.FillColor);
        Assert.Equal(OfficeColor.Lime, text.Color);
        Assert.Equal(10D, text.Font.Size, 3);
        Assert.Equal((100D, 12D), (item.Width, item.Height));
        Assert.Contains(
            rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(),
            visual => visual.Text == "EQUIVALENT");
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(container-type:inline-size)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(container-type:viewport)"));
    }

    [Fact]
    public void HtmlRendering_LaterContainerShorthandResetsAndOverridesItsLonghands() {
        var document = HtmlDocumentParser.ParseDocument("<style>#rule{container-name:old;container:new / inline-size}</style><section id='rule'></section><section id='inline' style='container-name:old;container:new / inline-size'></section>");
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> computed = HtmlComputedStyleEngine.Compute(document);

        foreach (string id in new[] { "rule", "inline" }) {
            HtmlComputedStyle style = computed[document.QuerySelector("#" + id)!];
            Assert.Equal("new", style.GetValue("container-name"));
            Assert.Equal("inline-size", style.GetValue("container-type"));
        }
    }

    [Fact]
    public void HtmlRendering_ContainerRangeQueriesAcceptCompactComparisonOperators() {
        const string html = "<style>@container (width>=300px){#minimum{background:red}}@container (300px<=width<400px){#range{background:red}}@container (width<300px){#outside{background:red}}</style><section style='width:360px;container-type:inline-size'><div id='minimum' style='width:20px;height:20px;background:blue'></div><div id='range' style='width:20px;height:20px;background:blue'></div><div id='outside' style='width:20px;height:20px;background:blue'></div></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 420D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#minimum").Shape.FillColor);
        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#range").Shape.FillColor);
        Assert.Equal(OfficeColor.Blue, FindShape(rendered, "div#outside").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_ContainerQueryThresholdUnitsUseTheQueriedContainer() {
        const string html = """
            <style>@container (width > 50cqw) { .item { background:red; } }</style>
            <section style="width:200px;container-type:inline-size"><div id="item" class="item">Container threshold</div></section>
            """;
        var options = new HtmlRenderOptions {
            ViewportWidth = 1000D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_NestedContainerUnitsUseTheNearestEligibleAncestorForGeometryAndFont() {
        const string html = "<style>@container (width:96px) and (height:36px) and (width > 4em){#item{background:red}}</style>"
            + "<section style='width:200px;height:100px;container-type:size'>"
            + "<div style='box-sizing:border-box;width:75cqw;height:80cqh;padding:5cqw 10cqw;border:1cqw solid black;min-width:60cqw;max-width:70cqw;min-height:50cqh;max-height:60cqh;font-size:10cqw;container-type:size'>"
            + "<div id='item' style='width:20px;height:20px;background:blue'></div></div></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 1000D, ViewportHeight = 500D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_ContainerUnitsSelectTheNearestEligibleAncestorPerAxis() {
        const string html = "<style>@container (width:50px) and (height:50px){#axis-item{background:red}}</style>"
            + "<section style='width:200px;height:100px;container-type:size'>"
            + "<div style='width:100px;container-type:inline-size'>"
            + "<div style='width:50cqw;height:50cqh;container-type:size'>"
            + "<div id='axis-item' style='width:20px;height:20px;background:blue'></div></div></div></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 1000D, ViewportHeight = 500D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#axis-item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_ContainerAspectRatioQueriesAcceptRatiosOnly() {
        const string html = "<style>@container (aspect-ratio:2){#ratio{background:red}}@container (aspect-ratio:2px){#length{background:red}}</style><section style='width:200px;height:100px;container-type:size'><div id='ratio' style='width:20px;height:20px;background:blue'></div><div id='length' style='width:20px;height:20px;background:blue'></div></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 300D, ViewportHeight = 200D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#ratio").Shape.FillColor);
        Assert.Equal(OfficeColor.Blue, FindShape(rendered, "div#length").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_SquareSizeContainerUsesPortraitOrientation() {
        const string html = "<style>@container (orientation:portrait){#item{background:red}}@container (orientation:landscape){#item{background:blue}}</style><section style='width:100px;height:100px;container-type:size'><div id='item' style='width:20px;height:20px;background:black'></div></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 200D, ViewportHeight = 200D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_GroupedContainerConditionsReevaluateLogicalOperators() {
        const string html = "<style>@container ((width > 100px) or (height > 100px)){#item{background:red}}</style><section style='width:120px;height:40px;container-type:size'><div id='item' style='width:20px;height:20px;background:blue'></div></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 300D, ViewportHeight = 200D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_BoxShorthandContainerUnitsUseTheActiveContainer() {
        const string html = "<section style='width:200px;container-type:inline-size'><div id='item' style='box-sizing:content-box;width:20px;height:20px;padding:10cqw;background:red'></div></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 800D, ViewportHeight = 400D });
        HtmlRenderShape item = FindShape(rendered, "div#item");

        Assert.Equal((60D, 60D), (item.Width, item.Height));
    }

    [Fact]
    public void HtmlRendering_ContainerQueryEmAndRemUnitsUseContainerAndRootFonts() {
        const string html = """
            <style>
              html { font-size:24px; }
              @container (width > 10em) { #em { background:red; } }
              @container (width > 7rem) { #rem { background:red; } }
              @container (width > 2em) { #relative { background:red; } }
            </style>
            <section style="width:180px;font-size:20px;container-type:inline-size">
              <div id="em" style="width:20px;height:20px;background:blue"></div>
              <div id="rem" style="width:20px;height:20px;background:blue"></div>
            </section>
            <div style="font-size:20px">
              <section style="width:70px;font-size:2em;container-type:inline-size">
                <div id="relative" style="width:20px;height:20px;background:blue"></div>
              </section>
            </div>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 300D });

        Assert.Equal(OfficeColor.Blue, FindShape(rendered, "div#em").Shape.FillColor);
        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#rem").Shape.FillColor);
        Assert.Equal(OfficeColor.Blue, FindShape(rendered, "div#relative").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_ContainerGeometryUsesComputedRootFontForRemLengths() {
        const string html = """
            <style>
              html { font-size:20px; }
              @container (width:120px) { #item { background:red; } }
            </style>
            <section style="box-sizing:border-box;width:10rem;padding:1rem;border:1rem solid black;container-type:inline-size">
              <div id="item" style="width:20px;height:20px;background:blue"></div>
            </section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 300D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_ContainerStyleQueriesCompareResolvedCustomPropertyValues() {
        const string html = """
            <style>@container theme style(--tone:red) { #item { background:red; } }</style>
            <section style="container-name:theme;--base:red;--tone:var(--base)">
              <div id="item" style="width:40px;height:20px;background:blue"></div>
            </section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_CustomPropertyNamesRemainCaseSensitiveInStyleQueriesAndVarResolution() {
        const string html = """
            <style>
              @container theme style(--Theme:red) { #exact { background:lime; } }
              @container theme style(--theme:red) { #wrong-case { background:red; } }
              @container theme style(--resolved:red) { #resolved { background:lime; } }
            </style>
            <section style="container-name:theme;--Theme:red">
              <div id="exact" style="width:20px;height:20px;background:blue"></div>
              <div id="wrong-case" style="width:20px;height:20px;background:blue"></div>
            </section>
            <section style="container-name:theme;--theme:red;--Theme:var(--theme);--resolved:var(--Theme)">
              <div id="resolved" style="width:20px;height:20px;background:blue"></div>
            </section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(OfficeColor.Lime, FindShape(rendered, "div#exact").Shape.FillColor);
        Assert.Equal(OfficeColor.Blue, FindShape(rendered, "div#wrong-case").Shape.FillColor);
        Assert.Equal(OfficeColor.Lime, FindShape(rendered, "div#resolved").Shape.FillColor);
    }

    [Theory]
    [InlineData("2em")]
    [InlineData("200%")]
    public void HtmlRendering_ContainerStyleQueriesResolveRelativeFontSizeFromTheParent(string relativeFontSize) {
        const string html = "<style>@container style(font-size:32px) { #item { background:red; } }</style>"
            + "<div style='font-size:16px'><section style='font-size:RELATIVE;container-name:theme'>"
            + "<div id='item' style='width:40px;height:20px;background:blue'></div></section></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html.Replace("RELATIVE", relativeFontSize), new HtmlRenderOptions());

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_NameOnlyContainerShorthandParticipatesInStyleQueries() {
        const string html = "<style>@container theme style(--tone:red) { #item { background:red; } }</style><section style='container:theme;--tone:red'><div id='item' style='width:40px;height:20px;background:blue'></div></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_ViewportUnitsReachTransformsAndGradientGeometry() {
        Assert.True(HtmlCssTransformParser.TryParse(
            "translate(10vw, 10vh)",
            "0 0",
            0D,
            0D,
            40D,
            20D,
            16D,
            16D,
            200D,
            100D,
            out OfficeTransform transform,
            out _));
        Assert.Equal(20D, transform.OffsetX, 6);
        Assert.Equal(10D, transform.OffsetY, 6);

        Assert.True(HtmlCssRadialGradientParser.TryParse(
            "radial-gradient(circle 10vw at 25vw 20vh, red, blue)",
            8,
            out HtmlCssRadialGradientDefinition? radial,
            out _));
        Assert.True(radial!.TryResolve(100D, 100D, 16D, 16D, 200D, 100D, out OfficeRadialGradient? gradient));
        Assert.Equal(0.5D, gradient!.EndX, 6);
        Assert.Equal(0.2D, gradient.EndY, 6);
        Assert.Equal(0.2D, gradient.EndRadiusX, 6);
    }

    [Fact]
    public void HtmlRendering_StyleQueriesDoNotTreatCustomPropertyNamesAsSizeFeatures() {
        const string html = """
            <style>@container theme style(--width: compact) { #item { background:red; } }</style>
            <section style="container-name:theme;--width:compact">
              <div id="item" style="width:40px;height:20px;background:blue"></div>
            </section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_ContainerQueriesMeasureTheContentBox() {
        const string html = """
            <style>@container (width:140px) { #item { background:red; } }</style>
            <section style="box-sizing:border-box;width:200px;padding:20px;border:10px solid black;container-type:inline-size">
              <div id="item" style="width:20px;height:20px;background:blue"></div>
            </section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 300D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_ContainerQueriesApplyMinAndMaxWidthConstraintsToTheContentBox() {
        const string html = """
            <style>
              @container (width:200px) { #maximum { background:red; } }
              @container (width:300px) { #minimum { background:red; } }
              @container (width:180px) { #border-box { background:red; } }
              @container (width:300px) { #conflict { background:red; } }
            </style>
            <section style="max-width:200px;container-type:inline-size">
              <div id="maximum" style="width:20px;height:20px;background:blue"></div>
            </section>
            <section style="width:100px;min-width:300px;container-type:inline-size">
              <div id="minimum" style="width:20px;height:20px;background:blue"></div>
            </section>
            <section style="box-sizing:border-box;max-width:240px;padding:20px;border:10px solid black;container-type:inline-size">
              <div id="border-box" style="width:20px;height:20px;background:blue"></div>
            </section>
            <section id="width-conflict" style="width:100px;min-width:300px;max-width:200px;container-type:inline-size;background:lime">
              <div id="conflict" style="width:20px;height:20px;background:blue"></div>
            </section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 800D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#maximum").Shape.FillColor);
        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#minimum").Shape.FillColor);
        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#border-box").Shape.FillColor);
        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#conflict").Shape.FillColor);
        Assert.Equal(300D, FindShape(rendered, "section#width-conflict").Width, 3);
    }

    [Fact]
    public void HtmlRendering_ContainerQueriesApplyMinAndMaxHeightConstraintsToTheContentBox() {
        const string html = """
            <style>
              @container (height:100px) { #maximum { background:red; } }
              @container (height:300px) { #minimum { background:red; } }
              @container (height:180px) { #border-box { background:red; } }
              @container (height:300px) { #conflict-height { background:red; } }
            </style>
            <section style="width:200px;height:200px;max-height:100px;container-type:size">
              <div id="maximum" style="width:20px;height:20px;background:blue"></div>
            </section>
            <section style="width:200px;height:100px;min-height:300px;container-type:size">
              <div id="minimum" style="width:20px;height:20px;background:blue"></div>
            </section>
            <section style="box-sizing:border-box;width:200px;height:300px;max-height:240px;padding:20px;border:10px solid black;container-type:size">
              <div id="border-box" style="width:20px;height:20px;background:blue"></div>
            </section>
            <section id="height-conflict" style="width:200px;height:100px;min-height:300px;max-height:200px;container-type:size;background:lime">
              <div id="conflict-height" style="width:20px;height:20px;background:blue"></div>
            </section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 800D, ViewportHeight = 1000D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#maximum").Shape.FillColor);
        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#minimum").Shape.FillColor);
        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#border-box").Shape.FillColor);
        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#conflict-height").Shape.FillColor);
        Assert.Equal(300D, FindShape(rendered, "section#height-conflict").Height, 3);
    }

    [Fact]
    public void HtmlRendering_AutoHeightSizeContainerUsesItsDefiniteMinimumForQueries() {
        const string html = """
            <style>
              @container (height:100px) { #content-box { background:red; } }
              @container (height:70px) { #border-box { background:red; } }
            </style>
            <section style="width:200px;min-height:100px;container-type:size">
              <div id="content-box" style="width:20px;height:20px;background:blue"></div>
            </section>
            <section style="box-sizing:border-box;width:200px;min-height:100px;padding:10px;border:5px solid black;container-type:size">
              <div id="border-box" style="width:20px;height:20px;background:blue"></div>
            </section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 500D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#content-box").Shape.FillColor);
        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#border-box").Shape.FillColor);
    }

    [Theory]
    [InlineData("(container-name:card)", true)]
    [InlineData("(container-name:-card)", true)]
    [InlineData("(container-name:\\31 23)", true)]
    [InlineData("(container-name:123)", false)]
    [InlineData("(container-name:-9card)", false)]
    [InlineData("(container:123 / size)", false)]
    public void HtmlRendering_ContainerSupportsRequiresCustomIdentifiers(string condition, bool expected) {
        Assert.Equal(expected, HtmlComputedStyleEngine.IsApplicableSupports(condition));
    }

    [Fact]
    public void HtmlRendering_SizeContainerPercentageHeightUsesItsDefiniteContainingBlock() {
        const string html = """
            <style>@container (height:100px) { #item { background:red; } }</style>
            <section style="height:200px">
              <div style="width:120px;height:50%;container-type:size">
                <div id="item" style="width:40px;height:20px;background:blue"></div>
              </div>
            </section>
            """;
        var options = new HtmlRenderOptions { ViewportHeight = 600D };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#item").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_SizeContainerDerivesItsContentHeightFromWidthAndAspectRatio() {
        const string html = """
            <style>
              @container (height:100px) { #content-box { background:red; } }
              @container (height:70px) { #border-box { background:red; } }
            </style>
            <section style="width:200px;aspect-ratio:2;container-type:size">
              <div id="content-box" style="width:20px;height:20px;background:blue"></div>
            </section>
            <section style="box-sizing:border-box;width:200px;aspect-ratio:2;padding:10px;border:5px solid black;container-type:size">
              <div id="border-box" style="width:20px;height:20px;background:blue"></div>
            </section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 500D });

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#content-box").Shape.FillColor);
        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#border-box").Shape.FillColor);
    }

    [Fact]
    public void HtmlRendering_CqhUsesTheSizeContainersContentBox() {
        const string html = "<section style='box-sizing:border-box;width:200px;height:200px;padding:20px;border:10px solid black;container-type:size'><div id='item' style='width:20px;height:50cqh;background:red'></div></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 300D, ViewportHeight = 400D });

        Assert.Equal(70D, FindShape(rendered, "div#item").Height, 3);
    }

    private static HtmlRenderShape FindShape(HtmlRenderDocument rendered, string source) =>
        Assert.Single(
            rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(),
            shape => shape.Source == source && shape.Shape.FillColor.HasValue);
}
