using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal readonly struct OfficeShadowLayer {
    internal OfficeShadowLayer(double strokeWidth, double expansion, double opacity, bool hasFill, bool hasStroke) {
        StrokeWidth = strokeWidth;
        Expansion = expansion;
        Opacity = opacity;
        HasFill = hasFill;
        HasStroke = hasStroke;
    }

    internal double StrokeWidth { get; }
    internal double Expansion { get; }
    internal double Opacity { get; }
    internal bool HasFill { get; }
    internal bool HasStroke { get; }
}

internal static class OfficeShadowLayerPlanner {
    private const int MinimumBlurLayers = 6;
    private const int MaximumBlurLayers = 16;

    internal static IReadOnlyList<OfficeShadowLayer> Create(
        double opacity,
        double blurRadius,
        double baseStrokeWidth,
        bool hasFill,
        bool hasStroke,
        bool canExpand,
        double minimumShapeDimension) {
        double clampedOpacity = Math.Max(0D, Math.Min(1D, opacity));
        double strokeWidth = Math.Max(0D, baseStrokeWidth);
        bool paintsFill = hasFill || !hasStroke;
        if (blurRadius <= 0D) {
            return new[] { new OfficeShadowLayer(strokeWidth, 0D, clampedOpacity, paintsFill, hasStroke) };
        }

        int layerCount = Math.Max(MinimumBlurLayers, Math.Min(MaximumBlurLayers, (int)Math.Ceiling(blurRadius / 2D)));
        if (canExpand && paintsFill) {
            // Expanded fill geometry represents the complete painted silhouette. A source
            // stroke extends half its width beyond the fill edge, including at the core.
            double silhouetteExpansion = hasStroke ? strokeWidth * 0.5D : 0D;
            // Reserve part of the opacity for the unexpanded core. This keeps a fully opaque
            // source opaque at its center without turning every expanded blur layer opaque.
            double blurCompositeOpacity = clampedOpacity * 0.5D;
            double totalWeight = layerCount * (layerCount + 1D) / 2D;
            var expandedLayers = new List<OfficeShadowLayer>(layerCount + 8);
            for (int index = layerCount; index >= 1; index--) {
                double weight = layerCount - index + 1D;
                double expandedLayerOpacity = blurCompositeOpacity <= 0D
                    ? 0D
                    : 1D - Math.Pow(1D - blurCompositeOpacity, weight / totalWeight);
                expandedLayers.Add(new OfficeShadowLayer(
                    0D,
                    silhouetteExpansion + blurRadius * index / layerCount,
                    expandedLayerOpacity,
                    hasFill: true,
                    hasStroke: false));
            }
            double coreOpacity = clampedOpacity >= 1D
                ? 1D
                : 1D - (1D - clampedOpacity) / Math.Max(0.000001D, 1D - blurCompositeOpacity);
            // A full-size opaque core gives a translated box shadow a hard rectangular
            // edge outside the source box. Spread that core inward so its edge fades.
            double maximumInset = Math.Min(blurRadius, Math.Max(0D, minimumShapeDimension * 0.45D));
            int coreLayerCount = maximumInset > 0.01D
                ? Math.Max(2, Math.Min(8, (int)Math.Ceiling(maximumInset / 4D)))
                : 1;
            // Keep the outer inset layers continuous at full opacity. The deepest
            // layer supplies the remainder, so the center still reaches the source opacity.
            double outerCoreCompositeOpacity = coreOpacity * 0.75D;
            double partialCoreOpacity = coreLayerCount == 1
                ? 0D
                : 1D - Math.Pow(1D - outerCoreCompositeOpacity, 1D / (coreLayerCount - 1D));
            double innerCoreOpacity = coreLayerCount == 1
                ? coreOpacity
                : 1D - (1D - coreOpacity) / (1D - outerCoreCompositeOpacity);
            for (int index = 0; index < coreLayerCount; index++) {
                double inset = coreLayerCount == 1 ? 0D : maximumInset * index / (coreLayerCount - 1D);
                expandedLayers.Add(new OfficeShadowLayer(
                    0D,
                    silhouetteExpansion - inset,
                    index == coreLayerCount - 1 ? innerCoreOpacity : partialCoreOpacity,
                    hasFill: true,
                    hasStroke: false));
            }
            return expandedLayers;
        }

        double ringOpacity = paintsFill ? 1D - Math.Sqrt(1D - clampedOpacity) : clampedOpacity;
        double layerOpacity = 1D - Math.Pow(1D - ringOpacity, 1D / layerCount);
        var layers = new List<OfficeShadowLayer>(layerCount + (paintsFill ? 1 : 0));
        for (int index = layerCount; index >= 1; index--) {
            double factor = index / (double)layerCount;
            layers.Add(new OfficeShadowLayer(
                Math.Max(0.5D, strokeWidth + blurRadius * 2D * factor),
                0D,
                layerOpacity,
                hasFill: false,
                hasStroke: true));
        }
        if (paintsFill) layers.Add(new OfficeShadowLayer(0D, 0D, clampedOpacity, hasFill: true, hasStroke: false));
        return layers;
    }

    internal static bool CanExpand(OfficeShape shape) => shape != null
        && !shape.Transform.HasValue
        && shape.ClipPath == null
        && (shape.Kind == OfficeShapeKind.Rectangle
            || shape.Kind == OfficeShapeKind.RoundedRectangle
            || shape.Kind == OfficeShapeKind.Ellipse);

    internal static OfficeShape CreateExpandedShape(OfficeShape shape, double expansion) {
        OfficeShape expanded = shape.Clone();
        expanded.Width += expansion * 2D;
        expanded.Height += expansion * 2D;
        if (expanded.Kind == OfficeShapeKind.RoundedRectangle) {
            expanded.CornerRadius = Math.Max(0D, Math.Min(
                expanded.CornerRadius + expansion,
                Math.Min(expanded.Width, expanded.Height) / 2D));
        }
        return expanded;
    }
}
