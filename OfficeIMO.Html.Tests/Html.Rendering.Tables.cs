using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlTableCell_StacksBlockDescendantsAndRetainsHeadingSemantics() {
        const string html = """
            <table style="width:320px;border-collapse:collapse">
              <tr><td style="padding:12px">
                <h1 style="margin:0 0 8px">Action Required</h1>
                <p style="margin:0 0 8px">Review the deployment package.</p>
                <div style="padding:8px;background:#fff4d6">Two checks need attention.</div>
              </td></tr>
            </table>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { ViewportWidth = 340D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderText heading = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Action Required");
        HtmlRenderText paragraph = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Review the deployment", StringComparison.Ordinal));
        HtmlRenderText notice = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Two checks", StringComparison.Ordinal));

        Assert.True(heading.Y < paragraph.Y);
        Assert.True(paragraph.Y < notice.Y);
        HtmlRenderHeading outline = Assert.Single(rendered.Headings);
        Assert.Equal("Action Required", outline.Text);
        Assert.Equal(1, outline.Level);

        string pdfText = PdfCore.PdfReadDocument.Open(
            HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions())).ExtractText();
        Assert.Contains("Action", pdfText, StringComparison.Ordinal);
        Assert.Contains("Review", pdfText, StringComparison.Ordinal);
        Assert.Contains("Two", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlTable_PaginatesInsideOneOversizedCellAtLineBoundaries() {
        const string html = """
            <table style="width:100px;border-collapse:collapse"><tbody><tr><td style="font-size:12px;line-height:20px">
              One<br>Two<br>Three<br>Four<br>Five<br>Six
            </td></tr></tbody></table>
            <div id="after-table" style="height:20px;background:#00ff00">After</div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 70D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.True(rendered.Pages.Count >= 2);
        Assert.All(rendered.Pages, page => Assert.True(page.Visuals.Count > 1));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Theory]
    [InlineData("top")]
    [InlineData("bottom")]
    public void HtmlTables_CaptionSidePaintsStyledCaptionAroundGridAcrossBackends(string side) {
        string html = "<body style='margin:0'><table id='table' style='width:80px;margin:0;caption-side:" + side + ";font-size:8px;line-height:10px'>"
            + "<caption id='caption' style='padding:2px;background:#ff0000'>CaptionPdf</caption>"
            + "<tr><td>CellPdf</td></tr></table></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderText caption = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "CaptionPdf");
        HtmlRenderText cell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "CellPdf");
        HtmlRenderShape captionBackground = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "caption#caption" && shape.Shape.FillColor == OfficeColor.Red);
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 50D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(80D, captionBackground.Width, 3);
        if (side == "top") Assert.True(caption.Y < cell.Y);
        else Assert.True(caption.Y > cell.Y);
        Assert.Contains("CaptionPdf", svg, StringComparison.Ordinal);
        Assert.Contains("CellPdf", svg, StringComparison.Ordinal);
        Assert.Contains("CaptionPdf", pdfText, StringComparison.Ordinal);
        Assert.Contains("CellPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlTables_EmptyGridRetainsItsCaption() {
        const string html = "<table style='width:60px;margin:0'><caption id='caption'>CaptionOnly</caption></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains("CaptionOnly", string.Concat(rendered.Text.Where(character => !char.IsWhiteSpace(character))), StringComparison.Ordinal);
        HtmlRenderSemanticGroup table = Assert.Single(rendered.Pages[0].Scene.OfType<HtmlRenderSemanticGroup>());
        Assert.Equal(HtmlRenderSemanticGroupRole.Table, table.Role);
        Assert.Contains(table.Visuals.OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Caption);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.EmptyTable);
    }

    [Fact]
    public void HtmlTables_AutoLayoutAllocatesColumnsFromIntrinsicCellContent() {
        const string html = "<table style='width:100px;margin:0;table-layout:auto;font-size:8px;line-height:10px'><tr>"
            + "<td id='wide' style='background:red'>WWWWWWWWWW</td><td id='narrow' style='background:blue'>i</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
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
        var options = new HtmlRenderOptions {
            ViewportWidth = 110D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape shared = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "table#conflict:collapsed-border-v-1-0");
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(110D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(OfficeColor.Red, shared.Shape.StrokeColor);
        Assert.Equal(5D, shared.Shape.StrokeWidth, 3);
        Assert.Contains("stroke=\"#FF0000\"", svg, StringComparison.Ordinal);
        Assert.Contains("LeftPdf", pdfText, StringComparison.Ordinal);
        Assert.Contains("RightPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlTables_CollapsedHiddenBorderSuppressesSharedCellEdge() {
        const string html = "<table id='hidden-conflict' style='width:100px;margin:0;table-layout:fixed;border-collapse:collapse'><tr>"
            + "<td style='border:1px solid black;border-right:5px solid red'>Left</td>"
            + "<td style='border:1px solid black;border-left:1px hidden blue'>Right</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 110D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "table#hidden-conflict:collapsed-border-v-1-0");
    }

    [Fact]
    public void HtmlTables_CollapsedBordersHonorCellRowAndTrackOriginPrecedence() {
        const string html = "<table id='origins' style='width:100px;margin:0;table-layout:fixed;border-collapse:collapse;border:2px solid purple'>"
            + "<colgroup style='border:2px solid orange'><col style='border-right:2px solid blue'></colgroup><colgroup><col></colgroup>"
            + "<tbody style='border:2px solid green'>"
            + "<tr style='border:2px solid blue'><td style='border:none;border-bottom:2px solid red'>A</td><td style='border:none'>B</td></tr>"
            + "<tr style='border:none'><td style='border:none'>C</td><td style='border:none'>D</td></tr>"
            + "</tbody></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 110D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderShape> shapes = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>().ToList();

        Assert.Equal(OfficeColor.Blue, Assert.Single(shapes, shape => shape.Source == "table#origins:collapsed-border-h-0-0").Shape.StrokeColor);
        Assert.Equal(OfficeColor.Red, Assert.Single(shapes, shape => shape.Source == "table#origins:collapsed-border-h-1-0").Shape.StrokeColor);
        Assert.Equal(OfficeColor.Blue, Assert.Single(shapes, shape => shape.Source == "table#origins:collapsed-border-v-1-0").Shape.StrokeColor);
        Assert.Equal(OfficeColor.Green, Assert.Single(shapes, shape => shape.Source == "table#origins:collapsed-border-h-2-1").Shape.StrokeColor);
    }

    [Fact]
    public void HtmlTables_CollapsedTableBorderPaintsOnlyResolvedOuterSegments() {
        const string html = "<table id='outer' style='width:100px;margin:0;table-layout:fixed;border-collapse:collapse;border:3px solid purple'>"
            + "<tr style='border:none;border-top:1px solid blue'><td style='border:none'>A</td><td style='border:none'>B</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 110D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderShape> shapes = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>().ToList();
        IReadOnlyList<HtmlRenderShape> collapsed = shapes
            .Where(shape => shape.Source != null && shape.Source.StartsWith("table#outer:collapsed-border-", StringComparison.Ordinal))
            .ToList();

        Assert.Equal(6, collapsed.Count);
        Assert.All(collapsed, shape => {
            Assert.Equal(OfficeColor.Purple, shape.Shape.StrokeColor);
            Assert.Equal(3D, shape.Shape.StrokeWidth, 3);
        });
        Assert.DoesNotContain(shapes, shape => shape.Source == "table#outer" && shape.Shape.StrokeColor == OfficeColor.Purple);
        Assert.DoesNotContain(collapsed, shape => shape.Source == "table#outer:collapsed-border-v-1-0");
    }

    [Fact]
    public void HtmlTables_InvalidCaptionSideUsesCatalogedTopFallbackAndSupportsTruth() {
        const string html = "<table id='table' style='caption-side:left;table-layout:balanced;border-collapse:merge;border-spacing:-2px;width:60px;margin:0'><caption>Caption</caption><tr><td>Cell</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);
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
