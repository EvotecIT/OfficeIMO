using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Measures layout overflow without treating shadows or glyph ink as box geometry.</summary>
internal static class HtmlRenderScrollableOverflow {
    internal static double MeasureRight(IReadOnlyList<HtmlRenderVisual> visuals, CancellationToken cancellationToken) =>
        Measure(visuals, OfficeTransform.Identity, new List<Clip>(), cancellationToken);

    private static double Measure(IReadOnlyList<HtmlRenderVisual> visuals, OfficeTransform transform,
        List<Clip> clips, CancellationToken cancellationToken) {
        double right = 0D;
        foreach (HtmlRenderVisual visual in visuals) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeTransform current = visual is HtmlRenderEffectGroup effect ? effect.Transform.Then(transform) : transform;
            IReadOnlyList<HtmlRenderVisual>? children = Children(visual);
            bool addedClip = false;
            if (visual is HtmlRenderClipGroup { IsViewportOverflow: false } rectangle && current.TryInvert(out OfficeTransform inverse)) {
                clips.Add(new Clip(inverse, rectangle.ClipX, rectangle.ClipY,
                    rectangle.ClipWidth, rectangle.ClipHeight, rectangle.ClipHorizontal, rectangle.ClipVertical));
                addedClip = true;
            } else if (visual is HtmlRenderPathClipGroup path && current.TryInvert(out OfficeTransform pathInverse)) {
                // Path bounds conservatively limit layout overflow; arbitrary path holes do not
                // change the underlying scrollable box geometry.
                clips.Add(new Clip(pathInverse, path.ClipX, path.ClipY, path.Width, path.Height, true, true));
                addedClip = true;
            }
            if (children != null) {
                right = Math.Max(right, Measure(children, current, clips, cancellationToken));
            } else if (visual is HtmlRenderLayoutBox || visual is HtmlRenderText { IsPaintOnly: false } || visual is HtmlRenderImage
                || visual is HtmlRenderDrawing || visual is HtmlRenderFormField) {
                var polygon = new List<OfficePoint> {
                    current.TransformPoint(new OfficePoint(visual.X, visual.Y)),
                    current.TransformPoint(new OfficePoint(visual.X + visual.Width, visual.Y)),
                    current.TransformPoint(new OfficePoint(visual.X + visual.Width, visual.Y + visual.Height)),
                    current.TransformPoint(new OfficePoint(visual.X, visual.Y + visual.Height))
                };
                foreach (Clip clip in clips) {
                    if (clip.Horizontal) {
                        polygon = Cut(polygon, clip, horizontal: true, upper: false);
                        polygon = Cut(polygon, clip, horizontal: true, upper: true);
                    }
                    if (clip.Vertical) {
                        polygon = Cut(polygon, clip, horizontal: false, upper: false);
                        polygon = Cut(polygon, clip, horizontal: false, upper: true);
                    }
                    if (polygon.Count == 0) break;
                }
                foreach (OfficePoint point in polygon) right = Math.Max(right, point.X);
            }
            if (addedClip) clips.RemoveAt(clips.Count - 1);
        }
        return right;
    }

    private static List<OfficePoint> Cut(List<OfficePoint> polygon, Clip clip, bool horizontal, bool upper) {
        var result = new List<OfficePoint>();
        if (polygon.Count == 0) return result;
        OfficePoint previous = polygon[polygon.Count - 1];
        double previousDistance = clip.Distance(previous, horizontal, upper);
        foreach (OfficePoint point in polygon) {
            double distance = clip.Distance(point, horizontal, upper);
            if ((distance >= 0D) != (previousDistance >= 0D)) {
                double ratio = previousDistance / (previousDistance - distance);
                result.Add(new OfficePoint(previous.X + (point.X - previous.X) * ratio,
                    previous.Y + (point.Y - previous.Y) * ratio));
            }
            if (distance >= 0D) result.Add(point);
            previous = point;
            previousDistance = distance;
        }
        return result;
    }

    /// <summary>Removes temporary geometry while preserving semantic ownership, clips and fonts.</summary>
    internal static IReadOnlyList<HtmlRenderVisual> RemoveLayoutBoxes(IReadOnlyList<HtmlRenderVisual> visuals,
        CancellationToken cancellationToken) {
        var result = new List<HtmlRenderVisual>(visuals.Count);
        bool changed = false;
        foreach (HtmlRenderVisual visual in visuals) {
            cancellationToken.ThrowIfCancellationRequested();
            if (visual is HtmlRenderLayoutBox) { changed = true; continue; }
            IReadOnlyList<HtmlRenderVisual>? children = Children(visual);
            IReadOnlyList<HtmlRenderVisual>? retained = children == null ? null : RemoveLayoutBoxes(children, cancellationToken);
            if (children != null && !ReferenceEquals(children, retained)) {
                changed = true;
                result.Add(visual switch {
                    HtmlRenderClipGroup clip => clip.ProjectPaint(retained!, 0D, 0D, visual.PaintOrder),
                    HtmlRenderPathClipGroup path => path.ProjectPaint(retained!, 0D, 0D, visual.PaintOrder),
                    HtmlRenderEffectGroup effect => effect.ProjectPaint(retained!, 0D, 0D, visual.PaintOrder),
                    HtmlRenderSemanticGroup semantic => semantic.ProjectPaint(retained!, 0D, 0D, visual.PaintOrder),
                    HtmlRenderLogicalTextGroup logical => logical.ProjectPaint(retained!, 0D, 0D, visual.PaintOrder, ownsLogicalText: true),
                    HtmlRenderLayoutRegion region => region.ProjectPaint(retained!, 0D, 0D, visual.PaintOrder),
                    HtmlRenderFormField form => form.ProjectPaint(retained!, 0D, 0D, visual.PaintOrder),
                    _ => visual
                });
            } else result.Add(visual);
        }
        return changed ? result : visuals;
    }

    private static IReadOnlyList<HtmlRenderVisual>? Children(HtmlRenderVisual visual) => visual switch {
        HtmlRenderClipGroup clip => clip.Visuals,
        HtmlRenderPathClipGroup path => path.Visuals,
        HtmlRenderEffectGroup effect => effect.Visuals,
        HtmlRenderSemanticGroup semantic => semantic.Visuals,
        HtmlRenderLogicalTextGroup logical => logical.Visuals,
        HtmlRenderLayoutRegion region => region.Visuals,
        HtmlRenderFormField form => form.Visuals,
        _ => null
    };

    private readonly struct Clip {
        private readonly OfficeTransform _inverse;
        private readonly double _x, _y, _width, _height;
        internal Clip(OfficeTransform inverse, double x, double y, double width, double height, bool horizontal, bool vertical) {
            _inverse = inverse; _x = x; _y = y; _width = width; _height = height;
            Horizontal = horizontal; Vertical = vertical;
        }
        internal bool Horizontal { get; }
        internal bool Vertical { get; }
        internal double Distance(OfficePoint point, bool horizontal, bool upper) {
            OfficePoint local = _inverse.TransformPoint(point);
            double value = horizontal ? local.X : local.Y;
            double minimum = horizontal ? _x : _y;
            double maximum = minimum + (horizontal ? _width : _height);
            return upper ? maximum - value : value - minimum;
        }
    }
}
