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
        bool canExpand) {
        double clampedOpacity = Math.Max(0D, Math.Min(1D, opacity));
        double strokeWidth = Math.Max(0D, baseStrokeWidth);
        bool paintsFill = hasFill || !hasStroke;
        if (blurRadius <= 0D) {
            return new[] { new OfficeShadowLayer(strokeWidth, 0D, clampedOpacity, paintsFill, hasStroke) };
        }

        int layerCount = Math.Max(MinimumBlurLayers, Math.Min(MaximumBlurLayers, (int)Math.Ceiling(blurRadius / 2D)));
        if (canExpand && paintsFill) {
            // Reserve part of the opacity for the unexpanded core. This keeps a fully opaque
            // source opaque at its center without turning every expanded blur layer opaque.
            double blurCompositeOpacity = clampedOpacity * 0.5D;
            double totalWeight = layerCount * (layerCount + 1D) / 2D;
            var expandedLayers = new List<OfficeShadowLayer>(layerCount + 1);
            for (int index = layerCount; index >= 1; index--) {
                double weight = layerCount - index + 1D;
                double expandedLayerOpacity = blurCompositeOpacity <= 0D
                    ? 0D
                    : 1D - Math.Pow(1D - blurCompositeOpacity, weight / totalWeight);
                expandedLayers.Add(new OfficeShadowLayer(
                    0D,
                    blurRadius * index / layerCount,
                    expandedLayerOpacity,
                    hasFill: true,
                    hasStroke: false));
            }
            double coreOpacity = clampedOpacity >= 1D
                ? 1D
                : 1D - (1D - clampedOpacity) / Math.Max(0.000001D, 1D - blurCompositeOpacity);
            expandedLayers.Add(new OfficeShadowLayer(0D, 0D, Math.Max(0D, Math.Min(1D, coreOpacity)), hasFill: true, hasStroke: false));
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
        if (paintsFill) layers.Add(new OfficeShadowLayer(0D, 0D, ringOpacity, hasFill: true, hasStroke: false));
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
            expanded.CornerRadius = Math.Min(
                expanded.CornerRadius + expansion,
                Math.Min(expanded.Width, expanded.Height) / 2D);
        }
        return expanded;
    }
}
