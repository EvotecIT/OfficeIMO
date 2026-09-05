using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static bool TryAddRasterizedDrawingEffect(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderDrawing visual,
        OfficeDrawing source,
        double rasterScale,
        PdfCore.PdfConversionReport conversionReport,
        CancellationToken cancellationToken) {
        if (!TryGetRasterizedDrawingEffectReason(source, out string effectReason)) return false;

        byte[] png = OfficeDrawingRasterRenderer.ToPng(source, new OfficeDrawingRasterRenderOptions {
            Scale = rasterScale,
            Background = OfficeColor.Transparent,
            CancellationToken = cancellationToken
        });
        PdfCore.PdfCanvasImageResource? effectImage = GetSharedPdfImageResource(png, "image/png");
        if (effectImage != null) {
            bool fragmentLink = IsFragmentLink(visual.LinkUri);
            canvas.ImageShared(
                effectImage,
                visual.X * PointsPerCssPixel,
                visual.Y * PointsPerCssPixel,
                visual.Width * PointsPerCssPixel,
                visual.Height * PointsPerCssPixel,
                linkUri: fragmentLink ? null : visual.LinkUri,
                linkContents: visual.LinkUri == null || fragmentLink ? null : visual.Source,
                alternativeText: visual.AlternativeText);
            if (fragmentLink) {
                canvas.LinkToNamedDestination(
                    MapNamedDestination(visual.LinkUri!.Substring(1)),
                    visual.X * PointsPerCssPixel,
                    visual.Y * PointsPerCssPixel,
                    visual.Width * PointsPerCssPixel,
                    visual.Height * PointsPerCssPixel,
                    visual.Source);
            }
        }
        conversionReport.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Html.Pdf",
            "HtmlPdfDrawingEffectRasterized",
            visual.Source ?? "html-drawing",
            "A managed drawing with " + effectReason + " was rasterized to preserve its compositing semantics in PDF output.",
            PdfCore.PdfConversionWarningSeverity.Warning,
            details: new Dictionary<string, string> {
                ["Effect"] = effectReason,
                ["Representation"] = "managed-png"
            }));
        return true;
    }

    private static bool TryGetRasterizedDrawingEffectReason(OfficeDrawing drawing, out string reason) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingEffectGroup effectGroup) {
                if (effectGroup.SoftMask != null) {
                    reason = "a vector soft mask";
                    return true;
                }
                if (effectGroup.BlendMode != OfficeBlendMode.Normal) {
                    reason = "the " + effectGroup.BlendMode + " blend mode";
                    return true;
                }
                if (TryGetRasterizedDrawingEffectReason(effectGroup.Drawing, out reason)) return true;
            } else if (element is OfficeDrawingGroup drawingGroup) {
                if (TryGetRasterizedDrawingEffectReason(drawingGroup.Drawing, out reason)) return true;
            } else if (element is OfficeDrawingTilingPattern tilingPattern
                       && TryGetRasterizedDrawingEffectReason(tilingPattern.Tile, out reason)) {
                return true;
            }
        }

        reason = string.Empty;
        return false;
    }
}
