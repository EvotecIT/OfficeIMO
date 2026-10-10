using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class HtmlEmptyInlineBlockBaselineTests {
    [Theory]
    [InlineData("<br>")]
    [InlineData("<span style='visibility:hidden'>X</span>")]
    public void UnpaintedInFlowLinesRetainTheirContentDerivedBaselineRoute(string children) {
        HtmlRenderDocument rendered = Render(Source("<div><span id='fill' style='display:inline-block;width:40px;height:240px;background:lime'>"
            + children + "</span></div>" + Following), Options());

        Assert.Equal(240D, Shape(rendered, "div#following").Y, 6);
    }

    [Theory]
    [InlineData(true, "baseline", 20D)]
    [InlineData(false, "baseline", 20D)]
    [InlineData(false, "top", 20D)]
    [InlineData(false, "baseline", 0D)]
    public void EmptyAtomicBoxIncludesTheParentStrutWithoutChangingItsHeight(bool table, string alignment, double lineHeight) {
        var options = Options();
        double descent = StrutDescent(options.Fonts, lineHeight);
        string box = "<span" + (table ? " style='height:20px'" : "") + "><div id='fill' style='display:inline-block;width:40px;background:lime;height:240px;" + (table ? "max-height:200%;" : "") + "vertical-align:" + alignment + "'></div></span>";
        string content = table ? "<table><tr><td style='height:120px'>" + box + "</td></tr></table>" : "<div>" + box + "</div>";
        HtmlRenderDocument rendered = Render(Source(content + Following, lineHeight), options);

        Assert.Equal(240D, Shape(rendered, "div#fill").Height, 6);
        Assert.Equal(0D, Shape(rendered, "div#fill").Y, 6);
        Assert.Equal(240D + (alignment == "baseline" ? descent : 0D), Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void WrappedAtomicLineBesideAFloatAdvancesByTheSameStrutAsPainting() {
        var options = Options();
        double descent = StrutDescent(options.Fonts, 20D);
        string html = Source("<div style='width:110px'><div style='float:left;width:20px;height:600px'></div>"
            + "<span id='first' style='display:inline-block;width:90px;height:240px;background:lime'></span>"
            + "<span id='second' style='display:inline-block;width:90px;height:240px;background:blue'></span></div>");
        HtmlRenderDocument rendered = Render(html, options);

        Assert.Equal(0D, Shape(rendered, "span#first").Y, 6);
        Assert.Equal(240D + descent, Shape(rendered, "span#second").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("", "Pinned")]
    [InlineData(" \n<!-- empty -->", "Missing Document Font,Pinned")]
    public void ActualEmptyInlineLayoutSuppliesTheBottomEdgeBaseline(string children, string families) {
        var options = Options();
        HtmlRenderDocument rendered = Render(Source("<div><span id='fill' style='display:inline-block;width:40px;height:240px;background:lime'>"
            + children + "</span></div>" + Following, families: families), options);

        Assert.Equal(240D + StrutDescent(options.Fonts, 20D), Shape(rendered, "div#following").Y, 6);
    }

    [Theory]
    [InlineData("Pinned")]
    [InlineData("Missing Document Font,Pinned")]
    public void PdfEmptyStrutUsesTheFirstAvailableRegisteredFace(string families) {
        var options = new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D), HonorCssPageRules = false };
        options.PdfOptions.RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily("Pinned", FontBytes()));
        Assert.True(options.PdfOptions.TryResolveNamedFontFace("Pinned", false, false, out var face));
        Assert.True(options.PdfOptions.TryGetNamedFontProgram(face, out var program));
        // The PDF program retains its existing integer metrics in 1000-em units.
        double ascent = program!.GetAscender(16D);
        double descent = Math.Max(0D, (20D + ascent + program.GetDescender(16D)) / 2D - ascent);
        string html = Source("<div><span id='fill' style='display:inline-block;width:40px;height:240px;background:lime'></span></div>" + Following, families: families);
        HtmlPdfRenderResult result = HtmlPdfRenderedConverter.Convert(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(240D, Shape(result.RenderResult!.Document, "span#fill").Height, 6);
        Assert.Equal(240D + descent, Shape(result.RenderResult.Document, "div#following").Y, 6);
        Assert.StartsWith("%PDF", System.Text.Encoding.ASCII.GetString(result.Document.ToBytes(), 0, 4));
        result.RenderResult.Document.RequireNoLoss();
    }

    private const string Following = "<div id='following' style='height:10px;background:blue'></div>";
    private static string Source(string content, double lineHeight = 20D, string families = "Pinned") =>
        "<!doctype html><style>html,body{margin:0;padding:0;font:16px/" + lineHeight.ToString(System.Globalization.CultureInfo.InvariantCulture)
        + "px " + families + "}table{margin:0;width:200px;border-spacing:0;table-layout:fixed}td{padding:0;vertical-align:top}</style>" + content;

    private static byte[] FontBytes() => File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf"));
    private static HtmlRenderOptions Options() {
        var options = new HtmlRenderOptions { ViewportWidth = 600D, ViewportHeight = 900D,
            PageSize = new OfficePageSize(600D / 96D, 900D / 96D), Margins = HtmlRenderMargins.All(0D),
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser, HonorCssPageRules = false };
        options.Fonts.Add("Pinned", FontBytes());
        return options;
    }

    private static double StrutDescent(OfficeFontFaceCollection fonts, double lineHeight) {
        IOfficeFontProgram program = fonts.ResolveForText(string.Empty, "Pinned", OfficeFontFaceDescriptor.Regular, 16D, out _)!;
        var baseline = Assert.IsAssignableFrom<IOfficeFontBaselineMetrics>(program);
        return Math.Max(0D, (lineHeight + program.LineHeight(16D)) / 2D - baseline.BaselineOffset(16D));
    }

    private static HtmlRenderDocument Render(string html, HtmlRenderOptions options) => HtmlRenderEngine.Execute(
        HtmlConversionDocument.Parse(html), HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: options)).Document;
    private static HtmlRenderShape Shape(HtmlRenderDocument document, string source) => Assert.Single(
        document.Pages.SelectMany(page => Enumerate(page.Scene)).OfType<HtmlRenderShape>(), shape => shape.Source == source && shape.Shape.FillColor.HasValue);
    private static IEnumerable<HtmlRenderVisual> Enumerate(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderSemanticGroup group => group.Visuals,
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                _ => null
            };
            if (children != null) foreach (HtmlRenderVisual child in Enumerate(children)) yield return child;
        }
    }
}
