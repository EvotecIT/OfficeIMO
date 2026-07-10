using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("top")]
    [InlineData("bottom")]
    public void HtmlTables_CaptionSidePaintsStyledCaptionAroundGridAcrossBackends(string side) {
        string html = "<body style='margin:0'><table id='table' style='width:80px;margin:0;caption-side:" + side + ";font-size:8px;line-height:10px'>"
            + "<caption id='caption' style='padding:2px;background:#ff0000'>CaptionPdf</caption>"
            + "<tr><td>CellPdf</td></tr></table></body>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 100D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderText caption = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "CaptionPdf");
        HtmlRenderText cell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "CellPdf");
        HtmlRenderShape captionBackground = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "caption#caption" && shape.Shape.FillColor == OfficeColor.Red);
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 50D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = html.SaveAsPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Load(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(80D, captionBackground.Width, 3);
        if (side == "top") Assert.True(caption.Y < cell.Y);
        else Assert.True(caption.Y > cell.Y);
        Assert.Contains("CaptionPdf", svg, StringComparison.Ordinal);
        Assert.Contains("CellPdf", svg, StringComparison.Ordinal);
        Assert.Contains("CaptionPdf", pdfText, StringComparison.Ordinal);
        Assert.Contains("CellPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlTables_EmptyGridRetainsItsCaption() {
        const string html = "<table style='width:60px;margin:0'><caption id='caption'>CaptionOnly</caption></table>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains("CaptionOnly", string.Concat(rendered.Text.Where(character => !char.IsWhiteSpace(character))), StringComparison.Ordinal);
        Assert.Contains(rendered.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.EmptyTable);
    }

    [Fact]
    public void HtmlTables_AutoLayoutAllocatesColumnsFromIntrinsicCellContent() {
        const string html = "<table style='width:100px;margin:0;table-layout:auto;font-size:8px;line-height:10px'><tr>"
            + "<td id='wide' style='background:red'>WWWWWWWWWW</td><td id='narrow' style='background:blue'>i</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape wide = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#wide" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderShape narrow = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#narrow" && shape.Shape.FillColor == OfficeColor.Blue);

        Assert.True(wide.Width > narrow.Width * 2D);
        Assert.Equal(100D, wide.Width + narrow.Width, 3);
        Assert.Equal(wide.Width, narrow.X, 3);
    }

    [Fact]
    public void HtmlTables_AutoLayoutIncludesIntrinsicReplacedImageWidth() {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(80, 10));
        string html = "<table style='width:100px;margin:0;table-layout:auto;font-size:8px;line-height:10px'><tr>"
            + "<td id='image-cell' style='background:red'><img src='data:image/png;base64," + imageData + "'></td>"
            + "<td id='text-cell' style='background:blue'>i</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape imageCell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#image-cell" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderShape textCell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#text-cell" && shape.Shape.FillColor == OfficeColor.Blue);

        Assert.True(imageCell.Width > textCell.Width * 4D);
        Assert.Equal(100D, imageCell.Width + textCell.Width, 3);
    }

    [Fact]
    public void HtmlTables_FixedLayoutHonorsColWidthsAndDistributesRemainder() {
        const string html = "<table style='width:100px;margin:0;table-layout:fixed;font-size:8px;line-height:10px'>"
            + "<colgroup><col style='width:70px'><col></colgroup><tr>"
            + "<td id='first' style='background:red'>A</td><td id='second' style='background:blue'>Long content does not resize this column</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape first = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#first" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderShape second = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#second" && shape.Shape.FillColor == OfficeColor.Blue);

        Assert.Equal(70D, first.Width, 3);
        Assert.Equal(30D, second.Width, 3);
        Assert.Equal(70D, second.X, 3);
    }

    [Fact]
    public void HtmlTables_SeparateBordersApplyHorizontalAndVerticalSpacing() {
        const string html = "<table style='width:100px;margin:0;table-layout:fixed;border-collapse:separate;border-spacing:4px 3px;font-size:8px;line-height:10px'>"
            + "<tr><td id='first' style='background:red'>A</td><td id='second' style='background:blue'>B</td></tr>"
            + "<tr><td id='third' style='background:lime'>C</td><td>D</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape first = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#first" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderShape second = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#second" && shape.Shape.FillColor == OfficeColor.Blue);
        HtmlRenderShape third = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#third" && shape.Shape.FillColor == OfficeColor.Lime);

        Assert.Equal((4D, 3D, 44D), (first.X, first.Y, first.Width));
        Assert.Equal((52D, 3D, 44D), (second.X, second.Y, second.Width));
        Assert.Equal(22D, third.Y, 3);
    }

    [Fact]
    public void HtmlTables_CollapsedBordersIgnoreBorderSpacingInGridGeometry() {
        const string html = "<table style='width:100px;margin:0;table-layout:fixed;border-collapse:collapse;border-spacing:10px;font-size:8px;line-height:10px'>"
            + "<tr><td id='first' style='background:red'>A</td><td id='second' style='background:blue'>B</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape first = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#first" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderShape second = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#second" && shape.Shape.FillColor == OfficeColor.Blue);

        Assert.Equal((0D, 0D, 50D), (first.X, first.Y, first.Width));
        Assert.Equal((50D, 0D, 50D), (second.X, second.Y, second.Width));
    }

    [Fact]
    public void HtmlTables_CollapsedBordersResolveSharedCellEdgeOnceAcrossBackends() {
        const string html = "<table id='conflict' style='width:100px;margin:0;table-layout:fixed;border-collapse:collapse;font-size:8px;line-height:10px'><tr>"
            + "<td style='border:1px solid black;border-right:5px solid red'>LeftPdf</td>"
            + "<td style='border:1px solid black;border-left:2px solid blue'>RightPdf</td></tr></table>";
        var options = new HtmlImageExportOptions {
            ViewportWidth = 110D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, options);
        HtmlRenderShape shared = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "table#conflict:collapsed-border-v-1-0");
        string svg = Encoding.UTF8.GetString(html.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = HtmlPdfSaveOptions.CreateRenderedProfile();
        pdfOptions.RenderOptions = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(110D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        byte[] pdf = html.SaveAsPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Load(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(OfficeColor.Red, shared.Shape.StrokeColor);
        Assert.Equal(5D, shared.Shape.StrokeWidth, 3);
        Assert.Contains("stroke=\"#FF0000\"", svg, StringComparison.Ordinal);
        Assert.Contains("LeftPdf", pdfText, StringComparison.Ordinal);
        Assert.Contains("RightPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(pdfOptions.ConversionReport.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlTables_CollapsedHiddenBorderSuppressesSharedCellEdge() {
        const string html = "<table id='hidden-conflict' style='width:100px;margin:0;table-layout:fixed;border-collapse:collapse'><tr>"
            + "<td style='border:1px solid black;border-right:5px solid red'>Left</td>"
            + "<td style='border:1px solid black;border-left:1px hidden blue'>Right</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 110D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "table#hidden-conflict:collapsed-border-v-1-0");
    }

    [Fact]
    public void HtmlTables_InvalidCaptionSideUsesCatalogedTopFallbackAndSupportsTruth() {
        const string html = "<table id='table' style='caption-side:left;table-layout:balanced;border-collapse:merge;border-spacing:-2px;width:60px;margin:0'><caption>Caption</caption><tr><td>Cell</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);
        HtmlRenderText caption = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Caption");
        HtmlRenderText cell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Cell");

        Assert.Equal("table#table", diagnostic.Source);
        Assert.Contains("caption-side=left", diagnostic.Detail, StringComparison.Ordinal);
        Assert.Contains("table-layout=balanced", diagnostic.Detail, StringComparison.Ordinal);
        Assert.Contains("border-collapse=merge", diagnostic.Detail, StringComparison.Ordinal);
        Assert.Contains("border-spacing=-2px", diagnostic.Detail, StringComparison.Ordinal);
        Assert.True(caption.Y < cell.Y);
        Assert.Contains(HtmlRenderDiagnosticCodes.TableValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.TableValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(caption-side:top)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(caption-side:bottom)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(caption-side:left)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(table-layout:auto)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(table-layout:fixed)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(table-layout:balanced)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-collapse:collapse)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-spacing:2px 4px)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border-spacing:-1px)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border-spacing:10%)"));
    }
}
