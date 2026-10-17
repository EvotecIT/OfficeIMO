using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Plans a bounded vector halo shared by raster, SVG and PDF shape writers.</summary>
internal static class OfficeGlowLayerPlanner {
    /// <summary>
    /// Returns detached strokes from widest to narrowest. Cumulative coverage increases toward
    /// the contour; only the innermost stroke can become opaque. Local transforms and clips survive.
    /// </summary>
    internal static IReadOnlyList<OfficeShape> CreateShapes(OfficeShape shape) {
        OfficeGlow? glow = shape.Glow;
        if (glow == null || glow.Radius <= 0D || glow.Opacity <= 0D || glow.Color.A == 0)
            return Array.Empty<OfficeShape>();

        int layerCount = (int)Math.Max(6D, Math.Min(16D, Math.Ceiling(glow.Radius / 2D)));
        double opacity = glow.Opacity * glow.Color.A / 255D;
        double baseStrokeWidth = Math.Max(0D, shape.StrokeWidth);
        OfficeShape halo = shape.Clone();
        halo.Glow = null;
        halo.Shadow = null;
        halo.FillColor = null;
        halo.FillGradient = null;
        halo.FillRadialGradient = null;
        halo.FillOpacity = 1D;
        halo.StrokeColor = OfficeColor.FromRgb(glow.Color.R, glow.Color.G, glow.Color.B);
        halo.StrokeGradient = null;
        halo.StrokeRadialGradient = null;
        halo.StrokeDashStyle = OfficeStrokeDashStyle.Solid;
        halo.SetStrokeDashArray(null);
        halo.StrokeHairline = false;
        halo.StrokeStartMarker = null;
        halo.StrokeEndMarker = null;
        halo.StrokeLineCap = OfficeStrokeLineCap.Round;
        halo.StrokeLineJoin = OfficeStrokeLineJoin.Round;

        var layers = new List<OfficeShape>(layerCount);
        double previousOpacity = 0D;
        for (int index = layerCount; index >= 1; index--) {
            double targetOpacity = opacity * (layerCount - index + 1D) / layerCount;
            OfficeShape layer = halo.Clone();
            layer.StrokeWidth = baseStrokeWidth + glow.Radius * (2D * index / layerCount);
            // Invert source-over composition so nested strokes reach the desired cumulative
            // opacity rather than applying the entire glow opacity independently at each radius.
            layer.StrokeOpacity = (targetOpacity - previousOpacity) / (1D - previousOpacity);
            layers.Add(layer);
            previousOpacity = targetOpacity;
        }
        return layers;
    }
}
