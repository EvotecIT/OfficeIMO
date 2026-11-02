using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlMargins_ParagraphAdjacentListNegativeTopMarginMovesItsFirstItem() {
        const string before = "<style>body{margin:0}p{margin:16px 0}ul{margin:16px 0;padding-left:40px}p+ul{margin-top:-8px}</style><p>Intro</p><ul><li>Item</li></ul>";
        const string withoutRule = "<style>body{margin:0}p{margin:16px 0}ul{margin:16px 0;padding-left:40px}</style><p>Intro</p><ul><li>Item</li></ul>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 320D, Margins = HtmlRenderMargins.All(0D) };
        HtmlConversionDocument document = HtmlConversionDocument.Parse(before);
        Assert.Equal("-8px", HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("ul")!].GetValue("margin-top"));
        HtmlRenderText withRule = Assert.Single(HtmlRenderTestDriver.Render(before, options).Pages[0].Visuals
            .OfType<HtmlRenderText>(), text => text.Text == "Item");
        HtmlRenderText unmodified = Assert.Single(HtmlRenderTestDriver.Render(withoutRule, options).Pages[0].Visuals
            .OfType<HtmlRenderText>(), text => text.Text == "Item");
        Assert.Equal(-8D, withRule.Y - unmodified.Y, 3);
    }

    [Fact]
    public void HtmlMargins_NegativeListMarginCollapsesThroughFirstItemsParagraph() {
        const string common = "<style>body{margin:0}p{margin:16px 0}ul{margin:16px 0;padding-left:40px}li{margin-bottom:8px}";
        const string content = "</style><div><p>Intro</p><ul><li><p>Item</p></li></ul></div>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 320D, Margins = HtmlRenderMargins.All(0D) };
        HtmlRenderText withRule = Assert.Single(HtmlRenderTestDriver.Render(common + "p+ul{margin-top:-8px}" + content, options)
            .Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Item");
        HtmlRenderText withoutRule = Assert.Single(HtmlRenderTestDriver.Render(common + content, options)
            .Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Item");
        Assert.Equal(-8D, withRule.Y - withoutRule.Y, 3);
    }

    [Fact]
    public void HtmlMargins_CollapseLongEmptyBlockRunsWithConstantState() {
        var html = new StringBuilder("<div style='margin:0'>");
        for (int index = 0; index < 2_000; index++) html.Append("<div style='margin:1px 0 2px'></div>");
        html.Append("<div id='tail' style='height:1px;margin:0;background:red'></div></div>");

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html.ToString(), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 20D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#tail");
    }

    [Fact]
    public void HtmlMargins_ParentAndFirstLastChildMarginsCollapseIntoExternalFlow() {
        const string html = "<div id='parent' style='width:50px;margin:10px 0 8px;background:#00ff00'>"
            + "<div id='child' style='width:40px;height:10px;margin:20px 0 12px;background:#ff0000;font-size:6px;line-height:8px'>ParentGap</div>"
            + "</div>"
            + "<div id='after-parent' style='width:40px;height:10px;margin:15px 0 0;background:#0000ff'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 55D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });
        HtmlRenderShape parent = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#parent");
        HtmlRenderShape child = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#child");
        HtmlRenderShape after = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#after-parent");

        Assert.Equal(20D, parent.Y, 3);
        Assert.Equal(10D, parent.Height, 3);
        Assert.Equal(parent.Y, child.Y, 3);
        Assert.Equal(45D, after.Y, 3);
    }

    [Fact]
    public void HtmlMargins_ParentPaddingPreventsChildMarginCollapse() {
        const string html = "<div id='padded-parent' style='width:50px;margin:0;padding:2px 0 3px;background:#00ff00'>"
            + "<div id='padded-child' style='width:40px;height:10px;margin:10px 0 12px;background:#ff0000'></div>"
            + "</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 55D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });
        HtmlRenderShape parent = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#padded-parent");
        HtmlRenderShape child = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#padded-child");

        Assert.Equal(0D, parent.Y, 3);
        Assert.Equal(37D, parent.Height, 3);
        Assert.Equal(12D, child.Y, 3);
    }

    [Fact]
    public void HtmlMargins_EmptyBlocksCollapseThroughAdjacentMarginSet() {
        const string html = "<div id='before-empty' style='width:20px;height:10px;margin:0 0 5px;background:#ff0000'></div>"
            + "<div id='empty' style='margin:10px 0 20px'></div>"
            + "<div id='after-empty' style='width:20px;height:10px;margin:15px 0 0;background:#0000ff'></div>"
            + "<div id='tail' style='width:20px;height:10px;margin:0;background:#00ff00'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 30D,
            ViewportHeight = 55D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });
        HtmlRenderShape before = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#before-empty");
        HtmlRenderShape after = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#after-empty");
        HtmlRenderShape tail = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#tail");

        Assert.Equal(0D, before.Y, 3);
        Assert.Equal(30D, after.Y, 3);
        Assert.Equal(40D, tail.Y, 3);
    }

    [Fact]
    public void HtmlContainerQueries_ContainerUnitsResolveInStylesheetPaddingShorthand() {
        const string html = "<style>.c{container-type:inline-size;width:200px}.p{padding:5cqw}</style><div class='c'><div class='p'>Text</div></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 220D, ViewportHeight = 60D, Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(10D, Assert.Single(DrawingTestTraversal.Elements(rendered.Pages[0].CreateDrawing()).OfType<OfficeDrawingText>()).X, 2);
    }

    [Theory]
    [InlineData("shape-outside:circle(50%)", "shape-outside")]
    [InlineData("text-emphasis:filled dot", "text-emphasis")]
    public void HtmlRendering_UnimplementedPaintPropertyIsReportedNotSilentlyDropped(string declaration, string property) {
        string html = "<style>.x{" + declaration + "}</style><p class='x'>Text</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions());

        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.CssPropertyUnsupported
            && item.Message.Contains(property, StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlContainerQueries_FlexItemQueriesItsLaidOutWidthNotItsParents() {
        const string html = "<style>.c{container-type:inline-size}.t{height:10px;background:#ff0000}"
            + "@container (min-width:150px){.t{background:#0000ff}}.u{height:10px;width:25cqw;background:#00ff00}</style>"
            + "<div style='display:flex;width:300px'><div class='c' style='flex:0 0 100px'><div id='narrow' class='t'></div><div id='unit' class='u'></div></div>"
            + "<div class='c' style='flex:1'><div id='wide' class='t'></div></div></div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 320D, ViewportHeight = 60D, Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Equal(OfficeColor.Red, FindShape(rendered, "div#narrow").Shape.FillColor);
        Assert.Equal(OfficeColor.Blue, FindShape(rendered, "div#wide").Shape.FillColor);
        Assert.Equal(25D, FindShape(rendered, "div#unit").Shape.Width, 2);
    }

    [Fact]
    public void HtmlGeneratedContent_AutoSizedAbsoluteMarkerSitsInItsRelativeParentsPadding() {
        const string html = "<style>li{position:relative;padding-left:20px;list-style:none}"
            + "li::before{content:'#';position:absolute;left:0}</style><ul style='margin:0;padding:0'><li>Text</li></ul>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 200D, ViewportHeight = 60D, Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        OfficeDrawingText[] texts = DrawingTestTraversal.Elements(rendered.Pages[0].CreateDrawing()).OfType<OfficeDrawingText>().ToArray();
        OfficeDrawingText marker = Assert.Single(texts, item => item.Text == "#");
        OfficeDrawingText body = Assert.Single(texts, item => item.Text == "Text");
        Assert.Equal(0D, marker.X, 2);
        Assert.Equal(20D, body.X, 2);
        Assert.Equal(body.Y, marker.Y, 2);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
    }

    [Fact]
    public void HtmlMargins_PdfKeepsBrowserBodyInsetUnlessAuthorResetsIt() {
        static double FirstWordX(string html) {
            var options = new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D), HonorCssPageRules = false };
            PdfCore.PdfLogicalTextBlock block = PdfCore.PdfDocumentReadResult.Load(
                HtmlConversionDocument.Parse(html).ToPdf(options).ToBytes()).Pages[0].TextBlocks[0];
            return block.XStart;
        }

        Assert.Equal(6D, FirstWordX("<div>Inset</div>"), 1);
        Assert.Equal(0D, FirstWordX("<style>body{margin:0}</style><div>Inset</div>"), 1);
    }

    [Fact]
    public void HtmlMargins_AdjacentBlockMarginsCollapseWithNegativeValuesAcrossBackends() {
        const string html = "<style>body{margin:0}</style><div id='container' style='width:50px;margin:0'>"
            + "<div id='first' style='width:20px;height:10px;margin:0 0 20px;background:#ff0000'></div>"
            + "<div id='second' style='width:20px;height:10px;margin:-5px 0 8px;background:#0000ff'></div>"
            + "<div id='third' style='width:40px;height:10px;margin:12px 0 0;background:#00ff00;font-size:6px;line-height:8px'>GapPdf</div>"
            + "</div>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 55D,
            ViewportHeight = 58D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape first = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#first");
        HtmlRenderShape second = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#second");
        HtmlRenderShape third = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "div#third");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlToPdfOptions pdfOptions = new HtmlToPdfOptions();
        pdfOptions = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(global::OfficeIMO.OfficeLength.Inches(55D / HtmlRenderOptions.CssPixelsPerInch), global::OfficeIMO.OfficeLength.Inches(58D / HtmlRenderOptions.CssPixelsPerInch)),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = (OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ToBytes();
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(0D, first.Y, 3);
        Assert.Equal(25D, second.Y, 3);
        Assert.Equal(47D, third.Y, 3);
        Assert.Equal(57D, third.Y + third.Height, 3);
        Assert.True(raster.GetPixel(5, 25).B > raster.GetPixel(5, 25).R);
        OfficeRasterImage svgRaster = RasterizeHtmlImageSvg(svg);
        Assert.Equal(OfficeColor.Transparent, svgRaster.GetPixel(5, 20));
        Assert.Equal(OfficeColor.Blue, svgRaster.GetPixel(5, 28));
        Assert.Equal(OfficeColor.Transparent, svgRaster.GetPixel(5, 40));
        Assert.Contains("GapPdf", pdfText, StringComparison.Ordinal);
        Assert.Single(PdfCore.PdfReadDocument.Open(pdf).Pages);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ConvertToPdfWithEvidence(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }
}
