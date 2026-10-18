using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(16D, 24D, 17D, 15D, 19D)]
    [InlineData(40D, 60D, 43D, 38D, 49D)]
    [InlineData(16D, 8D, 9D, 15D, 19D)]
    [InlineData(40D, 20D, 23D, 38D, 49D)]
    [InlineData(16D, 0D, 5D, 15D, 19D)]
    public void HtmlOrdinaryFontBox_UsedFaceKeepsItsHeightAcrossPositiveShortAndZeroLeading(
        double size, double lineHeight, double baseline, double ascent, double fontHeight) {
        // Independent marker-free native controls use the exact Baseline Sans
        // bytes:16px A15/D4 and40px A38/D11, with fixed CSS line heights.
        string html = OrdinaryFontBoxSource(size, lineHeight,
            "<span id='owner' style='background:lime'>Hgyp</span>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());
        HtmlRenderText text = Assert.Single(OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>());
        HtmlRenderShape owner = OrdinaryFontBoxShape(rendered, "span#owner");

        Assert.Equal(24D + baseline, text.Y + text.Font.Size, 6);
        Assert.Equal(24D + baseline - ascent, owner.Y, 6);
        Assert.Equal(fontHeight, owner.Height, 6);
        Assert.Equal(24D + lineHeight, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        AssertOrdinaryFontBoxPdf(html, "Hgyp", 24D + baseline);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("direct", false)]
    [InlineData("wrapped", false)]
    [InlineData("wrapped", true)]
    [InlineData("anonymous", false)]
    [InlineData("flow-root", false)]
    [InlineData("flow-root", true)]
    public void HtmlOrdinaryFontBox_ZeroLinesKeepPaintWithoutAddingFlowThroughWrappers(
        string layout, bool hidden) {
        string content = "<span id='owner' style='background:lime'><span style='visibility:"
            + (hidden ? "hidden" : "visible") + "'>Hgyp</span></span>";
        if (layout == "wrapped") content = "<div><div>" + content + "</div></div>";
        if (layout == "flow-root") content = "<div style='display:flow-root;background:#eee'>" + content + "</div>";
        if (layout == "anonymous") content += "<div style='height:0'></div>";
        string html = OrdinaryFontBoxSource(16D, 0D, content);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());

        Assert.Single(rendered.Pages);
        Assert.Equal(24D, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        Assert.Equal(19D, OrdinaryFontBoxShape(rendered, "span#owner").Height, 6);
        HtmlRenderText[] text = OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>().ToArray();
        if (hidden) Assert.Empty(text);
        else {
            Assert.Equal(29D, Assert.Single(text).Y + text[0].Font.Size, 6);
            AssertOrdinaryFontBoxPdf(html, "Hgyp", 29D);
        }
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(2D)]
    public void HtmlOrdinaryFontBox_ZeroFlowPaintSurvivesNonAtomicBlockInterruptions(double blockHeight) {
        string html = OrdinaryFontBoxSource(16D, 0D,
            "<span id='owner' style='background:lime'>Hgyp<div style='height:"
            + CssNumber(blockHeight) + "px'></div>TAIL</span>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());
        HtmlRenderText[] text = OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>().ToArray();

        Assert.Equal(2, text.Length);
        HtmlRenderText before = Assert.Single(text, item => item.Text == "Hgyp");
        HtmlRenderText after = Assert.Single(text, item => item.Text == "TAIL");
        Assert.Equal(29D, before.Y + before.Font.Size, 6);
        Assert.Equal(29D + blockHeight, after.Y + after.Font.Size, 6);
        Assert.Equal(24D + blockHeight, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(OrdinaryBaselineOptions()));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = pdf.GetPage(1);
        Assert.Equal("HgypTAIL", string.Concat(page.Letters.Select(letter => letter.Value)));
        for (int index = 0; index < page.Letters.Count; index++) {
            double expected = index < 4 ? 29D : 29D + blockHeight;
            Assert.InRange(Math.Abs((page.Height - page.Letters[index].StartBaseLine.Y) / 0.75D - expected),
                0D, 0.0005D / 0.75D + 0.000001D);
        }
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(0D, 5D, 1D, false)]
    [InlineData(0D, 5D, .5D, false)]
    [InlineData(8D, 9D, .5D, false)]
    [InlineData(24D, 17D, .5D, false)]
    [InlineData(0D, 5D, .5D, true)]
    public void HtmlOrdinaryFontBox_InterruptedFragmentsPaintInOrderInsideOneOwnerEffect(
        double lineHeight, double baseline, double opacity, bool nested) {
        string before = nested ? "<span style='background:blue;color:white'>Hgyp</span>" : "Hgyp";
        string html = OrdinaryFontBoxSource(16D, lineHeight,
            "<span id='owner' style='background:lime;opacity:" + CssNumber(opacity) + "'>" + before
            + "<div style='height:2px'></div>TAIL</span>");
        HtmlRenderOptions options = OrdinaryBaselineOptions();
        options.BackgroundColor = OfficeColor.Transparent;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderText[] text = OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>().ToArray();

        Assert.Equal("HgypTAIL", string.Concat(text.Select(item => item.Text)));
        Assert.Equal(24D + baseline, text[0].Y + text[0].Font.Size, 6);
        Assert.Equal(26D + lineHeight + baseline, text[1].Y + text[1].Font.Size, 6);
        Assert.Equal(26D + 2D * lineHeight, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        if (opacity < 1D) Assert.Single(OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderEffectGroup>(),
            group => group.Source == "span#owner" && group.Opacity == opacity);
        if (lineHeight < 16D) {
            // The later opaque fragment covers this earlier H stem. Native
            // marker-free PDF controls show lime here, with one owner opacity.
            int y = (int)(29D + lineHeight + baseline - 15D);
            OfficeColor pixel = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()).GetPixel(32, y);
            Assert.Equal((byte)0, pixel.R);
            Assert.Equal((byte)255, pixel.G);
            Assert.Equal((byte)0, pixel.B);
            Assert.InRange(pixel.A, (byte)(opacity * 255D), (byte)Math.Ceiling(opacity * 255D));
        }
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(OrdinaryBaselineOptions()));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal("HgypTAIL", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlOrdinaryFontBox_InterruptedOwnerRetainsBlockPaintAndTextInsideOneOpacityGroup() {
        string html = OrdinaryFontBoxSource(16D, 0D,
            "<span id='owner' style='background:lime;opacity:.5'>Hgyp"
            + "<div style='height:20px;line-height:20px;width:80px;margin-left:60px;"
            + "background:blue;border:1px solid red'>BLOCK</div>TAIL</span>");
        HtmlRenderOptions options = OrdinaryBaselineOptions();
        options.BackgroundColor = OfficeColor.Transparent;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        Assert.Single(OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderEffectGroup>(),
            group => group.Source == "span#owner" && group.Opacity == .5D);
        Assert.Equal("HgypBLOCKTAIL", string.Concat(OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>()
            .Select(text => text.Text)));
        Assert.Equal(46D, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()).GetPixel(160, 30);
        Assert.Equal((byte)0, pixel.R);
        Assert.Equal((byte)0, pixel.G);
        Assert.Equal((byte)255, pixel.B);
        Assert.InRange(pixel.A, (byte)127, (byte)128);
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(OrdinaryBaselineOptions()));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal("HgypBLOCKTAIL", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("opacity:.5", true)]
    [InlineData("transform:translateX(3px)", false)]
    public void HtmlOrdinaryFontBox_PlainInterruptedAncestorsRetainNestedPaintAndApplicableEffects(
        string effect, bool opacity) {
        string html = OrdinaryFontBoxSource(16D, 24D,
            "<span style='background:lime'><span style='background:yellow'>LEAD"
            + "<span id='grandchild' style='background:blue;" + effect + "'>WHOLE</span></span>"
            + "<div style='height:2px'></div>TAIL</span>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());
        HtmlRenderText[] text = OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>().ToArray();
        Assert.Equal(3, text.Length);
        Assert.Single(text, item => item.Text == "WHOLE");
        if (opacity) Assert.Single(OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderEffectGroup>(),
            group => group.Source == "span#grandchild");
        else Assert.DoesNotContain(OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderEffectGroup>(),
            group => group.Source == "span#grandchild");
        Assert.Equal(74D, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()).GetPixel(100, 27);
        if (opacity) {
            // Native blue at half opacity composites once over the plain
            // ancestor's yellow background at this point above the glyph ink.
            Assert.InRange(pixel.R, (byte)127, (byte)128);
            Assert.InRange(pixel.G, (byte)127, (byte)128);
            Assert.InRange(pixel.B, (byte)127, (byte)128);
        } else {
            Assert.Equal((byte)0, pixel.R);
            Assert.Equal((byte)0, pixel.G);
            Assert.Equal((byte)255, pixel.B);
        }
        Assert.Equal((byte)255, pixel.A);
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(OrdinaryBaselineOptions()));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        // Stacking layers can change content stream order; geometric reading
        // order proves every saved glyph survives exactly once on its line.
        Assert.Equal("LEADWHOLETAIL", string.Concat(pdf.GetPage(1).Letters
            .OrderByDescending(letter => letter.StartBaseLine.Y).ThenBy(letter => letter.StartBaseLine.X)
            .Select(letter => letter.Value)));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(.5D, "auto", true)]
    [InlineData(1D, "0", true)]
    [InlineData(.5D, "auto", false)]
    public void HtmlOrdinaryFontBox_InterruptedInlineOwnerKeepsNativeNegativeLayerPaintOrder(
        double opacity, string zIndex, bool background) {
        string html = OrdinaryFontBoxSource(16D, 0D,
            "<span id='owner' style='position:relative;opacity:" + CssNumber(opacity)
            + ";z-index:" + zIndex + (background ? ";background:lime" : "") + "'>MMMM"
            + "<div style='height:2px'></div>TAIL<span style='position:absolute;left:0;top:-20px;"
            + "width:80px;height:60px;background:red;z-index:-1'></span></span>");
        HtmlRenderOptions options = OrdinaryBaselineOptions();
        options.BackgroundColor = OfficeColor.Transparent;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing()).GetPixel(32, 17);
        Assert.Equal(background ? (byte)0 : (byte)255, pixel.R);
        Assert.Equal(background ? (byte)255 : (byte)0, pixel.G);
        Assert.Equal((byte)0, pixel.B);
        Assert.InRange(pixel.A, (byte)(opacity * 255D), (byte)Math.Ceiling(opacity * 255D));
        if (opacity < 1D) Assert.Single(OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderEffectGroup>(),
            group => group.Source == "span#owner" && group.Opacity == opacity);
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(OrdinaryBaselineOptions()));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal("MMMMTAIL", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("absolute")]
    [InlineData("fixed")]
    public void HtmlOrdinaryFontBox_PositionedOnlyMarkersRetainStaticAnchorsWithoutCreatingFlow(string position) {
        string html = OrdinaryFontBoxSource(16D, 24D,
            "<div style='position:relative'><span style='position:" + position + "'>X</span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());
        HtmlRenderText text = Assert.Single(OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>());
        Assert.Equal(24D, text.X, 6);
        Assert.Equal(41D, text.Y + text.Font.Size, 6);
        Assert.Equal(24D, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        AssertOrdinaryFontBoxPdf(html, "X", 41D);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlOrdinaryFontBox_PositionedMarkerFontDoesNotEnlargeTheSurroundingTextLine() {
        string Source(bool marker) => OrdinaryFontBoxSource(16D, 24D, "Hgyp" + (marker
            ? "<span style='position:absolute;left:100px;top:0;font:48px/72px Pinned'>X</span>" : ""));
        HtmlRenderDocument control = HtmlRenderTestDriver.Render(Source(false), OrdinaryBaselineOptions());
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(Source(true), OrdinaryBaselineOptions());
        HtmlRenderText expected = Assert.Single(OrdinaryFontBoxVisuals(control).OfType<HtmlRenderText>());
        HtmlRenderText actual = Assert.Single(OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>(), text => text.Text == "Hgyp");
        Assert.Equal(expected.Y + expected.Font.Size, actual.Y + actual.Font.Size, 6);
        Assert.Equal(OrdinaryFontBoxShape(control, "div#following").Y,
            OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        Assert.Equal(48D, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("top", 39D)]
    [InlineData("bottom", 29D)]
    public void HtmlOrdinaryFontBox_RemovingCollapsedMarginsPreservesTheZeroContentHeight(
        string side, double baseline) {
        string html = OrdinaryFontBoxSource(16D, 0D,
            "<div style='margin-" + side + ":10px'><span style='background:lime'>Hgyp</span></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());

        Assert.Equal(34D, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        HtmlRenderText text = Assert.Single(OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>());
        Assert.Equal(baseline, text.Y + text.Font.Size, 6);
        AssertOrdinaryFontBoxPdf(html, "Hgyp", baseline);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlOrdinaryFontBox_AnonymousZeroLineSeparatesAdjoiningMarginsEvenWhenHidden(bool hidden) {
        string html = OrdinaryFontBoxSource(16D, 0D,
            "<div id='before' style='height:10px;margin-bottom:20px;background:blue'></div>"
            + (hidden ? "<span style='visibility:hidden'>Hgyp</span>" : "Hgyp")
            + "<div id='after' style='height:10px;margin-top:30px;background:red'></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());

        Assert.Equal(24D, OrdinaryFontBoxShape(rendered, "div#before").Y, 6);
        Assert.Equal(84D, OrdinaryFontBoxShape(rendered, "div#after").Y, 6);
        Assert.Equal(94D, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        HtmlRenderText[] text = OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>().ToArray();
        if (hidden) Assert.Empty(text);
        else {
            Assert.Equal(59D, Assert.Single(text).Y + text[0].Font.Size, 6);
            AssertOrdinaryFontBoxPdf(html, "Hgyp", 59D);
        }
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlOrdinaryFontBox_OwnedZeroLineSeparatesAdjoiningMarginsEvenWhenHidden(bool hidden) {
        string html = OrdinaryFontBoxSource(16D, 0D,
            "<div id='before' style='height:10px;margin-bottom:20px;background:blue'></div>"
            + "<div style='visibility:" + (hidden ? "hidden" : "visible") + "'>Hgyp</div>"
            + "<div id='after' style='height:10px;margin-top:30px;background:red'></div>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());
        Assert.Equal(24D, OrdinaryFontBoxShape(rendered, "div#before").Y, 6);
        Assert.Equal(84D, OrdinaryFontBoxShape(rendered, "div#after").Y, 6);
        Assert.Equal(94D, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        HtmlRenderText[] text = OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>().ToArray();
        if (hidden) Assert.Empty(text);
        else {
            Assert.Equal(59D, Assert.Single(text).Y + text[0].Font.Size, 6);
            AssertOrdinaryFontBoxPdf(html, "Hgyp", 59D);
        }
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(16D, 24D, 40D, 60D, 43D, 60D)]
    [InlineData(16D, 8D, 40D, 20D, 23D, 22D)]
    [InlineData(40D, 20D, 16D, 8D, 23D, 22D)]
    public void HtmlOrdinaryFontBox_MixedSizesShareTheNativeBaselineAndUsedLineExtent(
        double parentSize, double parentLine, double childSize, double childLine, double baseline, double extent) {
        string html = OrdinaryFontBoxSource(parentSize, parentLine,
            "<span id='parent' style='background:lime'>H</span><span id='child' style='background:blue;font:"
            + CssNumber(childSize) + "px/" + CssNumber(childLine) + "px Pinned'>gyp</span>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());
        HtmlRenderText[] texts = OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>().ToArray();

        Assert.Equal(2, texts.Length);
        Assert.All(texts, text => Assert.Equal(24D + baseline, text.Y + text.Font.Size, 6));
        Assert.Equal(parentSize == 16D ? 19D : 49D, OrdinaryFontBoxShape(rendered, "span#parent").Height, 6);
        Assert.Equal(childSize == 16D ? 19D : 49D, OrdinaryFontBoxShape(rendered, "span#child").Height, 6);
        Assert.Equal(24D + extent, OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        AssertOrdinaryFontBoxPdf(html, "Hgyp", 24D + baseline);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(16D, 24D, 40D, 60D, 43D)]
    [InlineData(40D, 60D, 16D, 24D, 43D)]
    [InlineData(16D, 8D, 40D, 20D, 23D)]
    [InlineData(40D, 20D, 16D, 8D, 23D)]
    public void HtmlOrdinaryFontBox_NestedOwnersUseTheirOwnFontBoxesOnTheSharedBaseline(
        double outerSize, double outerLine, double innerSize, double innerLine, double baseline) {
        string html = OrdinaryFontBoxSource(outerSize, outerLine,
            "<span id='outer' style='padding:0 5px;border:1px solid black;background:lime'>H"
            + "<span id='inner' style='padding:0 3px;border:1px solid red;background:blue;font:"
            + CssNumber(innerSize) + "px/" + CssNumber(innerLine) + "px Pinned'>gyp</span></span>");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, OrdinaryBaselineOptions());
        HtmlRenderShape outer = OrdinaryFontBoxShape(rendered, "span#outer");
        HtmlRenderShape inner = OrdinaryFontBoxShape(rendered, "span#inner");

        Assert.Equal(outerSize == 16D ? 21D : 51D, outer.Height, 6);
        Assert.Equal(innerSize == 16D ? 21D : 51D, inner.Height, 6);
        Assert.Equal(24D + baseline - (outerSize == 16D ? 15D : 38D) - 1D, outer.Y, 6);
        Assert.Equal(24D + baseline - (innerSize == 16D ? 15D : 38D) - 1D, inner.Y, 6);
        Assert.All(OrdinaryFontBoxVisuals(rendered).OfType<HtmlRenderText>(),
            text => Assert.Equal(24D + baseline, text.Y + text.Font.Size, 6));
        AssertOrdinaryFontBoxPdf(html, "Hgyp", 24D + baseline);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlOrdinaryFontBox_DecorationDoesNotReplaceAnchorOrPositioningBounds() {
        string Source(bool decorated) => OrdinaryFontBoxSource(16D, 24D,
            "<span id='owner' style='position:relative;padding:0 5px;border:1px solid "
            + (decorated ? "black;background:lime" : "transparent") + "'>H"
            + "<a href='https://example.com/font-box' style='font:40px/60px Pinned'>gyp</a>"
            + "<span id='positioned' style='position:absolute;left:0;top:100%;width:6px;height:6px;background:red'></span>"
            + "</span>");
        HtmlRenderDocument control = HtmlRenderTestDriver.Render(Source(false), OrdinaryBaselineOptions());
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(Source(true), OrdinaryBaselineOptions());
        HtmlRenderShape positioned = OrdinaryFontBoxShape(rendered, "span#positioned");
        HtmlRenderShape controlPositioned = OrdinaryFontBoxShape(control, "span#positioned");

        Assert.Equal(21D, OrdinaryFontBoxShape(rendered, "span#owner").Height, 6);
        Assert.Equal(controlPositioned.X, positioned.X, 6);
        Assert.Equal(controlPositioned.Y, positioned.Y, 6);
        Assert.Equal(OrdinaryFontBoxShape(control, "div#following").Y,
            OrdinaryFontBoxShape(rendered, "div#following").Y, 6);
        PdfLinkAnnotation[] Links(string html) => PdfReadDocument.Open(HtmlConversionDocument.Parse(html)
            .ToPdfBytes(new HtmlToPdfOptions(OrdinaryBaselineOptions())))
            .Pages.SelectMany(page => page.GetLinkAnnotations()).ToArray();
        PdfLinkAnnotation[] expected = Links(Source(false));
        PdfLinkAnnotation[] actual = Links(Source(true));
        Assert.NotEmpty(expected);
        Assert.Equal(expected.Length, actual.Length);
        for (int index = 0; index < actual.Length; index++) {
            Assert.Equal(expected[index].Uri, actual[index].Uri);
            Assert.Equal(expected[index].X1, actual[index].X1, 6);
            Assert.Equal(expected[index].Y1, actual[index].Y1, 6);
            Assert.Equal(expected[index].X2, actual[index].X2, 6);
            Assert.Equal(expected[index].Y2, actual[index].Y2, 6);
        }
        rendered.RequireNoLoss();
    }

    private static string OrdinaryFontBoxSource(double size, double lineHeight, string content) =>
        "<!doctype html><style>body{margin:0;padding:24px;font:"
        + CssNumber(size) + "px/" + CssNumber(lineHeight) + "px Pinned}</style><div>" + content
        + "</div><div id='following' style='height:2px;background:red'></div>";

    private static string CssNumber(double value) => value.ToString(CultureInfo.InvariantCulture);

    private static IEnumerable<HtmlRenderVisual> OrdinaryFontBoxVisuals(HtmlRenderDocument document) =>
        document.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene));

    private static HtmlRenderShape OrdinaryFontBoxShape(HtmlRenderDocument document, string source) =>
        Assert.Single(OrdinaryFontBoxVisuals(document).OfType<HtmlRenderShape>(),
            shape => shape.Source == source && shape.Shape.FillColor.HasValue);

    private static void AssertOrdinaryFontBoxPdf(string html, string expectedText, double baseline) {
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(OrdinaryBaselineOptions()));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = pdf.GetPage(1);
        Assert.Equal(expectedText, string.Concat(page.Letters.Select(letter => letter.Value)));
        // The writer serializes point coordinates to three decimal places.
        // This is its exact half-unit bound, rather than the exploratory red
        // packet's broader diagnostic threshold.
        Assert.All(page.Letters, letter => Assert.InRange(
            Math.Abs((page.Height - letter.StartBaseLine.Y) / 0.75D - baseline),
            0D, 0.0005D / 0.75D + 0.000001D));
    }
}
