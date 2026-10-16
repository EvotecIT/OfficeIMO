using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    /// <summary>
    /// Paints a transform-only uniform enlargement directly on the destination grid.
    /// Source-over at full group opacity needs no isolated bitmap when no descendant
    /// blending depends on that isolation. Nearest images then sample their source
    /// once, instead of quantizing through a smaller intermediate drawing canvas.
    /// </summary>
    private static bool TryForwardScaledSourceOverGroup(OfficeRasterCanvas canvas,
        OfficeDrawingEffectGroup effect, EffectTextLayerPlan layer, double scale,
        IOfficeRasterImageCodec? imageCodec, long maximumRasterPixels, CancellationToken cancellationToken) {
        if (TryForwardTranslatedImage(canvas, effect, scale, imageCodec, maximumRasterPixels, cancellationToken)) return true;
        OfficeTransform transform = effect.Transform;
        if (layer.Interpolate || effect.Opacity != 1D || effect.BlendMode != OfficeBlendMode.Normal ||
            effect.SoftMask != null || effect.UnfilteredGeometryBounds.HasValue || layer.Left != 0D ||
            layer.Drawing.Width != effect.InnerDrawing.Width || layer.Drawing.Height != effect.InnerDrawing.Height ||
            transform.M11 <= 1D || transform.M11 != transform.M22 || transform.M12 != 0D || transform.M21 != 0D ||
            transform.OffsetX != 0D || transform.OffsetY != 0D) return false;
        double forwardedScale = scale * transform.M11;
        double width = Math.Ceiling(layer.Drawing.Width * forwardedScale * canvas.CoordinateScaleX);
        double height = Math.Ceiling(layer.Drawing.Height * forwardedScale * canvas.CoordinateScaleY);
        int maximumDimension = OfficeRasterImageEncoder.GetMaximumDimension(OfficeImageExportFormat.Png);
        // Keep the existing bounded renderer as the fallback for oversized or
        // inconclusive scenes; this path never allocates an unchecked surface.
        if (double.IsNaN(width) || double.IsInfinity(width) || double.IsNaN(height) || double.IsInfinity(height) ||
            width <= 0D || height <= 0D || width > maximumDimension || height > maximumDimension ||
            width * height > maximumRasterPixels || !HasOnlySourceOverCompositing(layer.Drawing, cancellationToken)) return false;
        using (canvas.PushClipRectangle(0D, 0D, layer.Drawing.Width * forwardedScale, layer.Drawing.Height * forwardedScale)) {
            // A fractional clip must cover the composited group once. Applying it
            // to each overlapping child changes both alpha and the visible color.
            if (!canvas.HasBinaryClipCoverage) return false;
            using var metricScope = canvas.PushFontMetricScale(forwardedScale);
            RenderElements(canvas, layer.Drawing.Elements, forwardedScale, imageCodec, maximumRasterPixels, cancellationToken);
        }
        return true;
    }

    /// <summary>Retains fractional placement on a single full-frame image until its final source sampling.</summary>
    private static bool TryForwardTranslatedImage(OfficeRasterCanvas canvas, OfficeDrawingEffectGroup effect,
        double scale, IOfficeRasterImageCodec? imageCodec, long maximumRasterPixels, CancellationToken cancellationToken) {
        OfficeTransform transform = effect.Transform;
        OfficeDrawing drawing = effect.InnerDrawing;
        if (effect.Opacity != 1D || effect.BlendMode != OfficeBlendMode.Normal || effect.SoftMask != null ||
            effect.UnfilteredGeometryBounds.HasValue || transform.M11 != 1D || transform.M22 != 1D ||
            transform.M12 != 0D || transform.M21 != 0D || drawing.Elements.Count != 1 ||
            drawing.Elements[0] is not OfficeDrawingImage image || image.Interpolate) return false;
        OfficeImageProjection projection = image.Projection;
        // Full-frame, unrotated image bounds are the layer's own implicit clip.
        // Crops remain on the projection and use the existing guarded image renderer.
        if (projection.HasTransform || projection.X != 0D || projection.Y != 0D ||
            projection.Width != drawing.Width || projection.Height != drawing.Height) return false;
        var placed = new OfficeDrawingImage(image.EncodedBytes, image.ContentType,
            projection.Translate(transform.OffsetX, transform.OffsetY), image.AlternativeText,
            image.Opacity, image.Interpolate, useDataSnapshot: true);
        RenderImage(canvas, placed, scale, imageCodec, maximumRasterPixels, cancellationToken);
        return true;
    }

    /// <summary>Proves that removing the outer isolation cannot change a descendant's blend backdrop.</summary>
    private static bool HasOnlySourceOverCompositing(OfficeDrawing drawing, CancellationToken cancellationToken) {
        const int maximumElements = 4096;
        int remaining = maximumElements;
        var pending = new Stack<(OfficeDrawing Drawing, int Depth)>();
        pending.Push((drawing, 0));
        while (pending.Count != 0) {
            (OfficeDrawing current, int depth) = pending.Pop();
            if (depth > 64) return false;
            foreach (OfficeDrawingElement element in current.Elements) {
                cancellationToken.ThrowIfCancellationRequested();
                if (--remaining < 0) return false;
                if (element is OfficeDrawingEffectGroup effect) {
                    if (effect.BlendMode != OfficeBlendMode.Normal || effect.SoftMask != null || effect.UnfilteredGeometryBounds.HasValue) return false;
                    pending.Push((effect.InnerDrawing, depth + 1));
                } else if (element is OfficeDrawingGroup group) {
                    pending.Push((group.InnerDrawing, depth + 1));
                } else if (element is OfficeDrawingTilingPattern pattern) {
                    pending.Push((pattern.InnerTile, depth + 1));
                } else if (element is not (OfficeDrawingShape or OfficeDrawingImage or OfficeDrawingImagePattern or OfficeDrawingText or OfficeDrawingRichText)) {
                    return false;
                }
            }
        }
        return true;
    }
}
