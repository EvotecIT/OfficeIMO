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

    private static HtmlRenderShape FindShape(HtmlRenderDocument rendered, string source) =>
        Assert.Single(
            rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(),
            shape => shape.Source == source && shape.Shape.FillColor.HasValue);
}
