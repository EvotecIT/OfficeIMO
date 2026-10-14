using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private readonly struct EffectTextLayerPlan {
        internal EffectTextLayerPlan(OfficeDrawing drawing, double left, double scaleX, double scaleY,
            bool interpolate, IReadOnlyList<(double Left, double Top, double Right, double Bottom)> paragraphInk) {
            Drawing = drawing; Left = left; ScaleX = scaleX; ScaleY = scaleY;
            Interpolate = interpolate; ParagraphInk = paragraphInk;
        }
        internal OfficeDrawing Drawing { get; }
        internal double Left { get; }
        internal double ScaleX { get; }
        internal double ScaleY { get; }
        internal bool Interpolate { get; }
        internal IReadOnlyList<(double Left, double Top, double Right, double Bottom)> ParagraphInk { get; }
        internal EffectTextLayerPlan WithInterpolation(bool interpolate) =>
            new EffectTextLayerPlan(Drawing, Left, ScaleX, ScaleY, interpolate, ParagraphInk);
    }

    private static EffectTextLayerPlan PlanEffectTextLayer(OfficeDrawingEffectGroup effect,
        OfficeRasterCanvas canvas, double scale, double parentX, double parentY,
        OfficeTransform? samplingPlacement = null) {
        // Nearest content retains its original grid. Measure that complete text
        // extent before deciding which images can affect the sampling policy.
        EffectTextLayerPlan original = CreateEffectTextLayer(effect.InnerDrawing, canvas,
            scale, scale, interpolate: false);
        SamplingInspectionContext inspection = new SamplingInspectionContext(canvas.CancellationToken);
        if (ContainsNearest(original)) return original;
        OfficeTransform transform = samplingPlacement.HasValue
            ? effect.Transform.Then(samplingPlacement.Value) : effect.Transform;
        (double axisX, double axisY) = GetEffectAxisScales(transform, parentX, parentY);
        double scaleX = scale * axisX, scaleY = scale * axisY;
        if (scaleX == scale && scaleY == scale) return original.WithInterpolation(true);
        // Layout uses one nominal font scale; painting and texel support retain
        // the independent axes so a narrow destination does not allocate a square.
        EffectTextLayerPlan sampled = CreateEffectTextLayer(effect.InnerDrawing, canvas,
            scaleX, scaleY, interpolate: true);
        // Density can change rounded overhang support. If that reveals nearest
        // content, conservatively return to its original grid and measurements.
        return ContainsNearest(sampled) ? original : sampled;

        bool ContainsNearest(EffectTextLayerPlan plan) {
            var surface = (Left: plan.Left, Top: 0D, Right: plan.Left + plan.Drawing.Width,
                Bottom: plan.Drawing.Height);
            return ContainsNonInterpolatedImage(effect.InnerDrawing, surface, inspection) ||
                (effect.SoftMask != null && ContainsVisibleNonInterpolatedImage(effect.SoftMask, surface, inspection));
        }
    }

    // An effect canvas isolates compositing, but is not a horizontal text viewport.
    // Keep each logical frame and the original vertical surface; expand only for
    // visible unwrapped paragraph ink measured with the layer's effective text profile.
    private static EffectTextLayerPlan CreateEffectTextLayer(OfficeDrawing drawing, OfficeRasterCanvas canvas,
        double scaleX, double scaleY, bool interpolate) {
        double scale = Math.Max(scaleX, scaleY);
        using var metricScope = canvas.PushFontMetricScale(scale);
        var paragraphInk = new List<(double Left, double Top, double Right, double Bottom)>(
            EffectParagraphInk(drawing, canvas, scale, scaleX / scale, scaleY / scale));
        double left = 0D, right = drawing.Width;
        foreach (var ink in paragraphInk) {
            left = Math.Min(left, ink.Left); right = Math.Max(right, ink.Right);
        }
        left = Math.Floor(left * scaleX) / scaleX;
        right = Math.Ceiling(right * scaleX) / scaleX;
        if (left == 0D && right == drawing.Width)
            return new EffectTextLayerPlan(drawing, left, scaleX, scaleY, interpolate, paragraphInk);
        var layer = new OfficeDrawing(right - left, drawing.Height);
        layer.AddDrawingForClippedRendering(drawing, -left, 0D, null);
        return new EffectTextLayerPlan(layer, left, scaleX, scaleY, interpolate, paragraphInk);
    }

    private static IEnumerable<(double Left, double Top, double Right, double Bottom)> EffectParagraphInk(
        OfficeDrawing drawing, OfficeRasterCanvas canvas, double scale, double coordinateX, double coordinateY,
        bool clipVertically = true, OfficeTransform? samplingPlacement = null) {
        canvas = canvas.WithDrawingTextProfile(drawing);
        foreach (OfficeDrawingElement element in drawing.Elements) {
            canvas.CancellationToken.ThrowIfCancellationRequested();
            if (element is OfficeDrawingRichText { WrapText: false } text && text.Paragraphs.Count > 0) {
                double width = (text.Width - text.Padding.Horizontal) * scale;
                double height = (text.Height - text.Padding.Vertical) * scale;
                if (width <= 0D || height <= 0D) continue;
                OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.CreateWithRasterMetrics(text, width, height, canvas, scale);
                foreach (var ink in OfficeDrawingTextLayout.PlacedRasterParagraphPaintBounds(layout, canvas,
                    (text.X + text.Padding.Left) * scale, (text.Y + text.Padding.Top) * scale,
                    height, text.VerticalAlignment)) {
                    var bounds = (Left: ink.Left / scale, Top: ink.Top / scale, Right: ink.Right / scale, Bottom: ink.Bottom / scale);
                    if (text.HasFrameTransform) bounds = TransformEffectInk(bounds, text.CreateFrameTransform().CreateDestinationTransform());
                    if (TryClipEffectInk(bounds, drawing.Height, clipVertically, out var visible)) yield return visible;
                }
            } else if (element is OfficeDrawingEffectGroup { Opacity: > 0D } effect) {
                if (!effect.Transform.TryInvert(out _)) continue;
                // Reuse the same measured layer plan as painting, retaining each
                // placed segment instead of combining hidden and visible lines.
                EffectTextLayerPlan child = PlanEffectTextLayer(effect, canvas, scale,
                    coordinateX, coordinateY, samplingPlacement);
                int pixelWidth = (int)Math.Min(int.MaxValue, Math.Max(1D, Math.Ceiling(child.Drawing.Width * child.ScaleX)));
                int pixelHeight = (int)Math.Min(int.MaxValue, Math.Max(1D, Math.Ceiling(child.Drawing.Height * child.ScaleY)));
                OfficeTransform samplingTransform = samplingPlacement.HasValue
                    ? effect.Transform.Then(samplingPlacement.Value) : effect.Transform;
                foreach (var ink in child.ParagraphInk) {
                    var bounds = TransformEffectInk(EffectRasterSamplingBounds(ink, samplingTransform, scale,
                        coordinateX, coordinateY, child.Left, child.ScaleX, child.ScaleY,
                        pixelWidth, pixelHeight, child.Interpolate), effect.Transform);
                    if (TryClipEffectInk(bounds, drawing.Height, clipVertically, out var visible)) yield return visible;
                }
            } else if (element is OfficeDrawingGroup group && group.ClipPath.Kind != OfficeClipPathKind.Empty) {
                OfficeTransform placement = OfficeTransform.Translate(group.X + group.ContentOffsetX, group.Y + group.ContentOffsetY);
                if (group.FrameTransform.HasValue) placement = placement.Then(group.FrameTransform.Value.CreateDestinationTransform());
                if (samplingPlacement.HasValue) placement = placement.Then(samplingPlacement.Value);
                foreach (var ink in EffectParagraphInk(group.InnerDrawing, canvas, scale, coordinateX, coordinateY,
                    clipVertically: false, samplingPlacement: placement)) {
                    var placed = TransformEffectInk(ink,
                        OfficeTransform.Translate(group.X + group.ContentOffsetX, group.Y + group.ContentOffsetY));
                    // The actual clip path remains on the layer; its rectangle is a
                    // conservative measurement boundary for nonrectangular paths.
                    if (!TryIntersectBounds(placed,
                        (group.X, group.Y, group.X + group.ClipPath.Width, group.Y + group.ClipPath.Height), out var clipped)) continue;
                    var bounds = group.FrameTransform.HasValue
                        ? TransformEffectInk(clipped, group.FrameTransform.Value.CreateDestinationTransform()) : clipped;
                    if (TryClipEffectInk(bounds, drawing.Height, clipVertically, out var visible)) yield return visible;
                }
            }
        }
    }

    private static (double Left, double Top, double Right, double Bottom) EffectRasterSamplingBounds(
        (double Left, double Top, double Right, double Bottom) ink, OfficeTransform transform, double scale,
        double coordinateX, double coordinateY, double layerLeft, double scaleX, double scaleY,
        int pixelWidth, int pixelHeight, bool interpolate) {
        // Match source texel coverage, area-prefilter buckets and the half-texel
        // bilinear kernel in the physical destination coordinates used by painting.
        double left = Math.Floor((ink.Left - layerLeft) * scaleX), right = Math.Ceiling((ink.Right - layerLeft) * scaleX);
        double top = Math.Floor(ink.Top * scaleY), bottom = Math.Ceiling(ink.Bottom * scaleY);
        OfficeTransform shifted = OfficeTransform.Translate(layerLeft, 0D).Then(transform);
        OfficeTransform deviceTransform = new OfficeTransform(
            shifted.M11 * scale / scaleX, shifted.M12 * scale / scaleX,
            shifted.M21 * scale / scaleY, shifted.M22 * scale / scaleY,
            shifted.OffsetX * scale, shifted.OffsetY * scale).Then(OfficeTransform.Scale(coordinateX, coordinateY));
        bool integerTranslation = deviceTransform.M11 == 1D && deviceTransform.M12 == 0D &&
            deviceTransform.M21 == 0D && deviceTransform.M22 == 1D &&
            deviceTransform.OffsetX == Math.Round(deviceTransform.OffsetX) && deviceTransform.OffsetY == Math.Round(deviceTransform.OffsetY);
        if (interpolate && !integerTranslation) {
            var step = OfficeRasterCanvas.AffineImageSamplingStep(pixelWidth, pixelHeight, deviceTransform);
            left = Math.Floor(left / step.X) * step.X - step.X / 2D;
            right = Math.Ceiling(right / step.X) * step.X + step.X / 2D;
            top = Math.Floor(top / step.Y) * step.Y - step.Y / 2D;
            bottom = Math.Ceiling(bottom / step.Y) * step.Y + step.Y / 2D;
        }
        return (layerLeft + Math.Max(0D, left) / scaleX, Math.Max(0D, top) / scaleY,
            layerLeft + Math.Min(pixelWidth, right) / scaleX, Math.Min(pixelHeight, bottom) / scaleY);
    }

    private static bool TryClipEffectInk((double Left, double Top, double Right, double Bottom) bounds,
        double height, bool clipVertically, out (double Left, double Top, double Right, double Bottom) visible) {
        if (clipVertically) return TryIntersectBounds(bounds, (bounds.Left, 0D, bounds.Right, height), out visible);
        visible = bounds;
        return true;
    }

    private static (double Left, double Top, double Right, double Bottom) TransformEffectInk(
        (double Left, double Top, double Right, double Bottom) bounds, OfficeTransform transform) =>
        transform.TransformRectangleBounds(bounds.Left, bounds.Top, bounds.Right - bounds.Left, bounds.Bottom - bounds.Top);
}
