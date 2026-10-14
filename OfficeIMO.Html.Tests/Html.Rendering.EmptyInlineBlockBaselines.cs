using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class HtmlEmptyInlineBlockBaselineTests {
    [Fact]
    public void NoWrapGroupRetainsItsEmptySiblingStrutWhenPlacedBelowAFloat() {
        var options = Options();
        HtmlRenderDocument rendered = Render(Source("<div style='width:110px'>"
            + "<i style='float:left;width:20px;height:120px'></i><span style='white-space:nowrap'>"
            + "<span style='line-height:100px'></span><span id='fill' style='display:inline-block;"
            + "width:100px;height:240px;background:lime'></span></span></div>" + Following), options);

        Assert.Equal(120D, Shape(rendered, "span#fill").Y, 6);
        Assert.Equal(360D + StrutDescent(options.Fonts, 100D), Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NoWrapGroupMovesItsEmptySiblingStrutWithItsAtomicContent(bool floated) {
        var options = Options();
        string floating = floated ? "<i style='float:left;width:20px;height:600px'></i>" : "";
        string content = "<div style='width:" + (floated ? "120" : "100") + "px'>" + floating
            + "<span id='first' style='display:inline-block;width:80px;height:240px;background:lime'></span>"
            + "<span style='white-space:nowrap'><span style='line-height:100px'></span>"
            + "<span id='second' style='display:inline-block;width:80px;height:240px;background:lime'></span></span></div>" + Following;
        HtmlRenderDocument rendered = Render(Source(content), options);

        Assert.Equal(0D, Shape(rendered, "span#first").Y, 6);
        Assert.Equal(240D + StrutDescent(options.Fonts, 20D), Shape(rendered, "span#second").Y, 6);
        Assert.Equal(480D + StrutDescent(options.Fonts, 20D) + StrutDescent(options.Fonts, 100D),
            Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FragmentDestinationMarkerDoesNotChangeTheParticipatingLineStruts(bool floated) {
        var options = Options();
        string floating = floated ? "<i style='float:left;width:20px;height:600px'></i>" : "";
        string content = "<div>" + floating + "<span id='anchor' style='line-height:100px'></span>"
            + "<span id='fill' style='display:inline-block;width:40px;height:240px;background:lime'></span></div>"
            + Following + "<div><a href='#anchor'>Jump</a></div>";
        HtmlRenderDocument rendered = Render(Source(content), options);

        Assert.Equal(240D + StrutDescent(options.Fonts, 100D), Shape(rendered, "div#following").Y, 6);
        Assert.Single(rendered.Pages.SelectMany(page => Enumerate(page.Scene)).OfType<HtmlRenderNamedDestination>(),
            destination => destination.Name == "anchor");
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("inline-block")]
    [InlineData("inline-flex")]
    [InlineData("inline-grid")]
    [InlineData("inline-table")]
    public void ReferencedInlineAtomicBoxOwnsOnePdfDestination(string display) {
        string html = Source("<div><span id='target' style='display:" + display
            + ";width:40px;height:240px;background:lime'></span></div>" + Following
            + "<div><a href='#target'>Jump</a></div>");
        HtmlPdfRenderRequestResult result = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options: Options()));

        Assert.Single(result.RenderResult.Document.Pages.SelectMany(page => Enumerate(page.Scene))
            .OfType<HtmlRenderNamedDestination>(), destination => destination.Name == "target");
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(result.ToBytes());
        Assert.Single(info.NamedDestinations, destination => destination.Name == "html-fragment:target");
        Assert.Contains("html-fragment:target", info.LinkDestinationNames);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" \n<!-- empty -->")]
    [InlineData("<span style='line-height:0'></span>")]
    public void EmptyInlineSiblingRetainsItsStrutOnTheAtomicLine(string children) {
        var options = Options();
        HtmlRenderDocument rendered = Render(Source("<div><span style='line-height:100px'>" + children
            + "</span><span id='fill' style='display:inline-block;width:40px;height:240px;background:lime'></span></div>" + Following), options);

        Assert.Equal(240D + StrutDescent(options.Fonts, 100D), Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void EmptyInlineStrutWithoutInFlowContentDoesNotCreateALine() {
        HtmlRenderDocument emptyBlock = Render(Source("<div></div>" + Following), Options());
        HtmlRenderDocument emptyInline = Render(Source("<div><span style='line-height:100px'></span></div>" + Following), Options());

        Assert.Equal(Shape(emptyBlock, "div#following").Y, Shape(emptyInline, "div#following").Y, 6);
    }

    [Theory]
    [InlineData(true, 80D, 0D, 100D)]
    [InlineData(false, 80D, 0D, 100D)]
    [InlineData(false, 100D, 120D, 20D)]
    public void EmptySiblingStrutFollowsFloatLinePlacement(bool precedesFloat, double width, double expectedTop, double lineHeight) {
        var options = Options();
        const string strut = "<span style='line-height:100px'></span>";
        const string floating = "<div style='float:left;width:20px;height:120px'></div>";
        string fill = "<span id='fill' style='display:inline-block;width:" + width + "px;height:240px;background:lime'></span>";
        HtmlRenderDocument rendered = Render(Source("<div style='width:110px'>"
            + (precedesFloat ? strut + floating : floating + strut) + fill + "</div>" + Following), options);

        Assert.Equal(expectedTop, Shape(rendered, "span#fill").Y, 6);
        Assert.Equal(expectedTop + 240D + StrutDescent(options.Fonts, lineHeight), Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(1D)]
    public void PlainInlineWrapperContributesItsStrutIndependentlyOfHorizontalPadding(double padding) {
        var options = Options();
        string html = Source("<div><span style='line-height:100px;padding-left:" + padding
            + "px'><span id='fill' style='display:inline-block;width:40px;height:240px;background:lime'></span></span></div>" + Following);
        HtmlRenderDocument rendered = Render(html, options);

        Assert.Equal(240D, Shape(rendered, "span#fill").Height, 6);
        Assert.Equal(240D + StrutDescent(options.Fonts, 100D), Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void OuterInlineWrapperStrutSurvivesANestedShorterLineHeight() {
        var options = Options();
        string html = Source("<div><span style='line-height:100px'><span style='line-height:0'>"
            + "<span id='fill' style='display:inline-block;width:40px;height:240px;background:lime'></span></span></span></div>" + Following);
        HtmlRenderDocument rendered = Render(html, options);

        Assert.Equal(240D + StrutDescent(options.Fonts, 100D), Shape(rendered, "div#following").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(1D)]
    public void WrappedAtomicLinesBesideAFloatRetainThePlainWrapperStrut(double padding) {
        var options = Options();
        string html = Source("<div style='width:110px'><div style='float:left;width:20px;height:600px'></div>"
            + "<span style='line-height:100px;padding-left:" + padding + "px'>"
            + "<span id='first' style='display:inline-block;width:80px;height:240px;background:lime'></span> "
            + "<span id='second' style='display:inline-block;width:80px;height:240px;background:blue'></span></span></div>");
        HtmlRenderDocument rendered = Render(html, options);

        Assert.Equal(240D + StrutDescent(options.Fonts, 100D), Shape(rendered, "span#second").Y, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("font-size:20px")]
    [InlineData("vertical-align:super")]
    public void UnqualifiedWrapperMetricsRetainTheSameRouteWithOrWithoutDecoration(string metrics) {
        HtmlRenderDocument Plain(bool padded) => Render(Source("<div><span style='" + metrics
            + (padded ? ";padding-left:1px" : "")
            + "'><span id='fill' style='display:inline-block;width:40px;height:240px;background:lime'></span></span></div>" + Following), Options());
        HtmlRenderDocument plain = Plain(false);
        HtmlRenderDocument padded = Plain(true);

        Assert.Equal(240D, Shape(plain, "span#fill").Height, 6);
        Assert.Equal(Shape(padded, "div#following").Y, Shape(plain, "div#following").Y, 6);
    }

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
