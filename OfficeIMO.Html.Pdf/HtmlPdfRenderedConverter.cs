using OfficeIMO.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static class HtmlPdfRenderedConverter {
    private const double PointsPerCssPixel = 72D / HtmlRenderOptions.CssPixelsPerInch;

    internal static PdfCore.PdfDocument Convert(string html, HtmlPdfSaveOptions options) {
        HtmlRenderOptions renderOptions = options.RenderOptions?.Clone() ?? new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged
        };
        options.RenderOptions = renderOptions;
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(html, renderOptions);
        return CreatePdf(rendered, options);
    }

    internal static async Task<PdfCore.PdfDocument> ConvertAsync(string html, HtmlPdfSaveOptions options, CancellationToken cancellationToken) {
        HtmlRenderOptions renderOptions = options.RenderOptions?.Clone() ?? new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged
        };
        options.RenderOptions = renderOptions;
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(html, renderOptions, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return CreatePdf(rendered, options);
    }

    private static PdfCore.PdfDocument CreatePdf(HtmlRenderDocument rendered, HtmlPdfSaveOptions options) {
        options.RenderDiagnostics = rendered.Diagnostics.Clone();

        PdfCore.PdfDocument pdf = PdfCore.PdfDocument.Create();
        if (options.RenderedFontFamily != null) {
            pdf.UseFontFamily(options.RenderedFontFamily);
        }

        pdf.UseTextFallbacks(options.RenderedTextFallbacks)
            .UseTextShaping(options.RenderedTextShapingMode, options.RenderedTextShapingProvider);
        foreach (HtmlRenderPage renderedPage in rendered.Pages) {
            double pageWidth = renderedPage.Width * PointsPerCssPixel;
            double pageHeight = renderedPage.Height * PointsPerCssPixel;
            pdf.Page(page => page
                .Size(pageWidth, pageHeight)
                .Margin(0D)
                .Canvas(canvas => AddPageVisuals(canvas, renderedPage)));
        }

        return pdf;
    }

    private static void AddPageVisuals(PdfCore.PdfPageCanvas canvas, HtmlRenderPage page) {
        foreach (HtmlRenderVisual visual in page.Visuals.OrderBy(item => item.PaintOrder)) {
            if (visual is HtmlRenderShape shape) {
                AddShape(canvas, shape);
            } else if (visual is HtmlRenderText text) {
                AddText(canvas, text);
            } else if (visual is HtmlRenderImage image) {
                AddImage(canvas, image);
            }
        }
    }

    private static void AddShape(PdfCore.PdfPageCanvas canvas, HtmlRenderShape visual) {
        var drawing = new OfficeDrawing(visual.Width, visual.Height);
        drawing.AddShape(visual.Shape.Clone(), 0D, 0D);
        canvas.Drawing(
            drawing,
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            linkUri: visual.LinkUri,
            linkContents: visual.LinkUri == null ? null : visual.Source);
    }

    private static void AddText(PdfCore.PdfPageCanvas canvas, HtmlRenderText visual) {
        if (visual.Text.Length == 0) return;
        string? link = string.IsNullOrWhiteSpace(visual.Text) ? null : visual.LinkUri;
        var run = new PdfCore.TextRun(
            visual.Text,
            bold: visual.Font.IsBold,
            underline: visual.Font.IsUnderline,
            color: PdfCore.PdfColor.FromOfficeColorOrNull(visual.Color),
            italic: visual.Font.IsItalic,
            strike: visual.Font.IsStrikethrough,
            fontSize: visual.Font.Size * PointsPerCssPixel,
            font: MapFont(visual.Font.FamilyName),
            linkUri: link,
            linkContents: link == null ? null : visual.Text);
        canvas.Text(
            new[] { run },
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            PdfCore.PdfColor.FromOfficeColorOrNull(visual.Color),
            MapAlignment(visual.Alignment),
            visual.Font.Size * PointsPerCssPixel,
            visual.LineHeight * PointsPerCssPixel);
    }

    private static void AddImage(PdfCore.PdfPageCanvas canvas, HtmlRenderImage visual) {
        canvas.Image(
            visual.Bytes,
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            linkUri: visual.LinkUri,
            linkContents: visual.LinkUri == null ? null : visual.Source,
            alternativeText: visual.AlternativeText);
    }

    private static PdfCore.PdfStandardFont MapFont(string familyName) {
        string normalized = familyName ?? string.Empty;
        if (normalized.IndexOf("times", StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf("serif", StringComparison.OrdinalIgnoreCase) >= 0) {
            return PdfCore.PdfStandardFont.TimesRoman;
        }

        if (normalized.IndexOf("courier", StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf("consolas", StringComparison.OrdinalIgnoreCase) >= 0
            || normalized.IndexOf("mono", StringComparison.OrdinalIgnoreCase) >= 0) {
            return PdfCore.PdfStandardFont.Courier;
        }

        return PdfCore.PdfStandardFont.Helvetica;
    }

    private static PdfCore.PdfAlign MapAlignment(OfficeTextAlignment alignment) {
        if (alignment == OfficeTextAlignment.Center) return PdfCore.PdfAlign.Center;
        if (alignment == OfficeTextAlignment.Right) return PdfCore.PdfAlign.Right;
        if (alignment == OfficeTextAlignment.Justify) return PdfCore.PdfAlign.Justify;
        return PdfCore.PdfAlign.Left;
    }
}
