using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
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
    public void HtmlTables_InvalidCaptionSideUsesCatalogedTopFallbackAndSupportsTruth() {
        const string html = "<table id='table' style='caption-side:left;width:60px;margin:0'><caption>Caption</caption><tr><td>Cell</td></tr></table>";

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
        Assert.True(caption.Y < cell.Y);
        Assert.Contains(HtmlRenderDiagnosticCodes.TableValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.TableValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(caption-side:top)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(caption-side:bottom)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(caption-side:left)"));
    }
}
