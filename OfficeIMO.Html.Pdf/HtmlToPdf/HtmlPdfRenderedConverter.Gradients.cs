using System.Globalization;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    /// <summary>
    /// Rasterizes a gradient only when native PDF shading cannot carry its stops.
    /// </summary>
    /// <remarks>
    /// The writer keeps translucent stops as vector shading with a luminosity soft mask and
    /// approximates CSS premultiplied interpolation with intermediate color samples, so a
    /// transparent stop does not leave a grey fringe. Stop sets beyond the shading function's
    /// bounds fall back to an image at the export's raster density.
    /// </remarks>
    private static bool TryAddRasterizedGradient(
        PdfCore.PdfPageCanvas canvas,
        HtmlRenderShape visual,
        OfficeDrawing drawing,
        PdfImageResourceCache imageResources,
        PdfCore.PdfConversionReport conversionReport,
        CancellationToken cancellationToken,
        bool suppressLink) {
        // Inspect the detached descriptor already retained by this drawing.
        OfficeShape shape = drawing.Shapes[0].Shape;
        if (CanWriteNativeGradient(shape)) return false;

        cancellationToken.ThrowIfCancellationRequested();
        byte[] png = OfficeDrawingRasterRenderer.ToPng(drawing, new OfficeDrawingRasterRenderOptions {
            Scale = imageResources.EffectRasterScale,
            Background = OfficeColor.Transparent,
            MaximumRasterPixels = imageResources.MaximumPixels,
            CancellationToken = cancellationToken
        });
        PdfCore.PdfCanvasImageResource? image = imageResources.GetOrCreate(png, "image/png");
        if (image != null) {
            canvas.ImageShared(
                image,
                visual.X * PointsPerCssPixel,
                visual.Y * PointsPerCssPixel,
                visual.Width * PointsPerCssPixel,
                visual.Height * PointsPerCssPixel,
                linkUri: suppressLink ? null : visual.LinkUri,
                linkContents: suppressLink || visual.LinkUri == null ? null : visual.Source);
        }
        conversionReport.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Html.Pdf",
            "HtmlPdfTranslucentGradientRasterized",
            visual.Source ?? "html-gradient",
            "A CSS gradient exceeded the stops native PDF shading can carry and was rasterized.",
            PdfCore.PdfConversionWarningSeverity.Warning,
            details: new Dictionary<string, string> {
                ["Representation"] = "managed-png",
                ["Interpolation"] = "premultiplied-alpha",
                ["Dpi"] = (imageResources.EffectRasterScale * HtmlRenderOptions.CssPixelsPerInch).ToString("0.###", CultureInfo.InvariantCulture)
            }));
        return true;
    }

    private static bool CanWriteNativeGradient(OfficeShape shape) {
        if (shape.FillRadialGradient is OfficeRadialGradient radial) {
            return PdfCore.PdfVisualResourceDictionaryBuilder.CanWriteGradient(radial.Stops, radial.ColorInterpolation, radial.InterpolateAlphaSeparately);
        }
        return shape.FillGradient is not OfficeLinearGradient linear
            || PdfCore.PdfVisualResourceDictionaryBuilder.CanWriteGradient(linear.Stops, linear.ColorInterpolation, linear.InterpolateAlphaSeparately);
    }
}
