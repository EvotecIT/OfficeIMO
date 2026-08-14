using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static bool TryAddTranslucentGradient(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderShape visual,
        OfficeDrawing drawing,
        PdfCore.PdfConversionReport conversionReport,
        CancellationToken cancellationToken) {
        IReadOnlyList<OfficeGradientStop>? stops = visual.Shape.FillRadialGradient?.Stops ?? visual.Shape.FillGradient?.Stops;
        if (stops == null || stops.All(stop => stop.Color.A == byte.MaxValue)) return false;

        cancellationToken.ThrowIfCancellationRequested();
        byte[] png = OfficeDrawingRasterRenderer.ToPng(drawing, new OfficeDrawingRasterRenderOptions {
            Background = OfficeColor.Transparent,
            CancellationToken = cancellationToken
        });
        PdfCore.PdfCanvasImageResource? image = GetSharedPdfImageResource(png, "image/png");
        if (image != null) {
            canvas.ImageShared(
                image,
                visual.X * PointsPerCssPixel,
                visual.Y * PointsPerCssPixel,
                visual.Width * PointsPerCssPixel,
                visual.Height * PointsPerCssPixel,
                linkUri: visual.LinkUri,
                linkContents: visual.LinkUri == null ? null : visual.Source);
        }
        conversionReport.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Html.Pdf",
            "HtmlPdfTranslucentGradientRasterized",
            visual.Source ?? "html-gradient",
            "A translucent CSS gradient was rasterized with premultiplied-alpha interpolation to preserve its compositing semantics in PDF output.",
            PdfCore.PdfConversionWarningSeverity.Information,
            details: new Dictionary<string, string> {
                ["Representation"] = "managed-png",
                ["Interpolation"] = "premultiplied-alpha"
            }));
        return true;
    }
}
