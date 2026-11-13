using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    private static void AppendDrawingContent(StringBuilder builder, OfficeDrawing drawing,
        IOfficeRasterImageCodec? imageCodec, string idPrefix, ref int gradientId, ref int clipPathId,
        CancellationToken cancellationToken, SvgTilingExpansionBudget tilingExpansionBudget,
        OfficeDrawingRasterRenderOptions rasterOptions, OfficeRasterCanvas textMetrics, bool textAsPaths,
        System.Collections.Generic.ISet<string> requiredFontFamilies) {
        if (drawing.PostCompositionColorTransform == null) {
            AppendElements(builder, drawing.Elements, imageCodec, idPrefix, ref gradientId, ref clipPathId,
                cancellationToken, tilingExpansionBudget, rasterOptions, textMetrics, textAsPaths, requiredFontFamilies);
            return;
        }

        // Portable SVG has no destination ICC transform for a composited group.
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
            Scale = rasterOptions.Scale,
            MaximumRasterPixels = rasterOptions.MaximumRasterPixels,
            ImageCodec = imageCodec,
            DiagnosticSink = rasterOptions.DiagnosticSink,
            DiagnosticSource = rasterOptions.DiagnosticSource,
            ThrowOnImageDecodeFailure = true,
            InheritedFonts = textMetrics.Fonts,
            TextShapingProvider = textMetrics.TextShapingProvider,
            TextShapingLanguage = textMetrics.TextShapingLanguage,
            CancellationToken = cancellationToken
        });
        cancellationToken.ThrowIfCancellationRequested();
        builder.Append("<image x=\"0\" y=\"0\" width=\"").Append(Format(drawing.Width))
            .Append("\" height=\"").Append(Format(drawing.Height))
            .Append("\" preserveAspectRatio=\"none\" href=\"data:image/png;base64,");
        long remainingCharacters = builder.MaxCapacity - (long)builder.Length - 3L;
        if (remainingCharacters < 4L)
            throw new ArgumentOutOfRangeException(nameof(builder), "The encoded SVG exceeds its configured character ceiling.");
        byte[] png;
        try {
            png = OfficeRasterImageEncoder.Encode(raster, OfficeImageExportFormat.Png, null,
                remainingCharacters / 4L * 3L, cancellationToken);
        } catch (OfficeImageExportBatchLimitException) {
            // The public export limit describes the enclosing SVG, not this derived PNG allowance.
            throw new ArgumentOutOfRangeException(nameof(builder), "The encoded SVG exceeds its configured character ceiling.");
        }
        EnsureBase64PayloadFits(builder, png.LongLength);
        OfficeSvgImageRenderer.AppendBase64(builder, png, cancellationToken);
        builder.Append("\"/>");
    }

    // sqrt(||A||1 * ||A||infinity) bounds the placement density without overflowing squared coefficients.
    private static OfficeDrawingRasterRenderOptions TransformRasterOptions(
        OfficeDrawingRasterRenderOptions options, OfficeTransform transform) {
        double largest = Math.Max(Math.Max(Math.Abs(transform.M11), Math.Abs(transform.M12)),
            Math.Max(Math.Abs(transform.M21), Math.Abs(transform.M22)));
        if (largest == 0D) return options;
        double a = Math.Abs(transform.M11) / largest, b = Math.Abs(transform.M12) / largest;
        double c = Math.Abs(transform.M21) / largest, d = Math.Abs(transform.M22) / largest;
        double density = largest * Math.Sqrt(Math.Max(a + b, c + d) * Math.Max(a + c, b + d));
        return new OfficeDrawingRasterRenderOptions {
            Scale = options.Scale * density,
            MaximumRasterPixels = options.MaximumRasterPixels,
            DiagnosticSink = options.DiagnosticSink,
            DiagnosticSource = options.DiagnosticSource
        };
    }
}
