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
            TextShapingProvider = textMetrics.TextShapingProvider,
            TextShapingLanguage = textMetrics.TextShapingLanguage,
            CancellationToken = cancellationToken
        });
        byte[] png = OfficePngWriter.Encode(raster, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        builder.Append("<image x=\"0\" y=\"0\" width=\"").Append(Format(drawing.Width))
            .Append("\" height=\"").Append(Format(drawing.Height))
            .Append("\" preserveAspectRatio=\"none\" href=\"data:image/png;base64,")
            .Append(Convert.ToBase64String(png)).Append("\"/>");
    }
}
