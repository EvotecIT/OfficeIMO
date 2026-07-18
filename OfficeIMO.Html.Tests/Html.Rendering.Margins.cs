using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlMargins_ParentAndFirstLastChildMarginsCollapseIntoExternalFlow() {
        const string html = "<div id='parent' style='width:50px;margin:10px 0 8px;background:#00ff00'>"
            + "<div id='child' style='width:40px;height:10px;margin:20px 0 12px;background:#ff0000;font-size:6px;line-height:8px'>ParentGap</div>"
            + "</div>"
            + "<div id='after-parent' style='width:40px;height:10px;margin:15px 0 0;background:#0000ff'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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
    public void HtmlMargins_AdjacentBlockMarginsCollapseWithNegativeValuesAcrossBackends() {
        const string html = "<div id='container' style='width:50px;margin:0'>"
            + "<div id='first' style='width:20px;height:10px;margin:0 0 20px;background:#ff0000'></div>"
            + "<div id='second' style='width:20px;height:10px;margin:-5px 0 8px;background:#0000ff'></div>"
            + "<div id='third' style='width:40px;height:10px;margin:12px 0 0;background:#00ff00;font-size:6px;line-height:8px'>GapPdf</div>"
            + "</div>";
        var options = new HtmlRenderOptions {
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
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(55D / HtmlRenderOptions.CssPixelsPerInch, 58D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(0D, first.Y, 3);
        Assert.Equal(25D, second.Y, 3);
        Assert.Equal(47D, third.Y, 3);
        Assert.Equal(57D, third.Y + third.Height, 3);
        Assert.True(raster.GetPixel(5, 25).B > raster.GetPixel(5, 25).R);
        Assert.Contains("y=\"25\"", svg, StringComparison.Ordinal);
        Assert.Contains("GapPdf", pdfText, StringComparison.Ordinal);
        Assert.Single(PdfCore.PdfReadDocument.Open(pdf).Pages);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }
}
