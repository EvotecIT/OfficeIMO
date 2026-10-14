using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Creates bounded, cancellable page projections while preserving only intersecting scene nodes.</summary>
internal sealed class HtmlRenderVisualProjector {
    private readonly int _maximumVisuals;
    private readonly CancellationToken _cancellationToken;
    private readonly bool _navigationAlreadyPageOwned;
    private readonly HashSet<HtmlRenderLogicalTextGroup> _ownedLogicalText = new();
    private readonly HashSet<HtmlRenderText> _ownedText = new();
    private readonly HashSet<HtmlRenderVisual> _ownedNavigation = new();
    private readonly List<ProjectionClip> _clips = new();
    private int _materializedVisuals;

    internal HtmlRenderVisualProjector(int maximumVisuals, CancellationToken cancellationToken, bool navigationAlreadyPageOwned = false) {
        if (maximumVisuals <= 0) throw new ArgumentOutOfRangeException(nameof(maximumVisuals));
        _maximumVisuals = maximumVisuals;
        _cancellationToken = cancellationToken;
        _navigationAlreadyPageOwned = navigationAlreadyPageOwned;
    }

    internal IReadOnlyList<HtmlRenderVisual> Project(
        IEnumerable<HtmlRenderVisual> visuals,
        double sourceX,
        double sourceY,
        double width,
        double height,
        double outputX,
        double outputY) {
        if (visuals == null) throw new ArgumentNullException(nameof(visuals));
        var projected = new List<HtmlRenderVisual>();
        double right = sourceX + width;
        double bottom = sourceY + height;
        double offsetX = outputX - sourceX;
        double offsetY = outputY - sourceY;
        foreach (HtmlRenderVisual visual in visuals.OrderBy(item => item.PaintOrder)) {
            HtmlRenderVisual? result = Project(visual, sourceX, sourceY, right, bottom, offsetX, offsetY,
                projected.Count, OfficeTransform.Identity);
            if (result != null) projected.Add(result);
        }
        return projected.AsReadOnly();
    }

    private HtmlRenderVisual? Project(
        HtmlRenderVisual visual,
        double left,
        double top,
        double right,
        double bottom,
        double offsetX,
        double offsetY,
        int paintOrder,
        OfficeTransform transform) {
        _cancellationToken.ThrowIfCancellationRequested();

        if (visual is HtmlRenderClipGroup clipGroup) {
            return ProjectClipped(clipGroup.Visuals,
                new ProjectionClip(clipGroup.ClipX, clipGroup.ClipY, clipGroup.ClipWidth, clipGroup.ClipHeight,
                    clipGroup.ClipHorizontal, clipGroup.ClipVertical, transform),
                children => clipGroup.ProjectPaint(children, offsetX, offsetY, paintOrder));
        }
        if (visual is HtmlRenderPathClipGroup pathClipGroup) {
            return ProjectClipped(pathClipGroup.Visuals,
                new ProjectionClip(pathClipGroup.ClipX, pathClipGroup.ClipY, pathClipGroup.Width, pathClipGroup.Height,
                    true, true, transform),
                children => pathClipGroup.ProjectPaint(children, offsetX, offsetY, paintOrder));
        }
        if (visual is HtmlRenderEffectGroup effectGroup) {
            return ProjectGroup(effectGroup.Visuals, children => effectGroup.ProjectPaint(children, offsetX, offsetY, paintOrder),
                effectGroup.Transform.Then(transform));
        }
        if (visual is HtmlRenderSemanticGroup semanticGroup) {
            return ProjectGroup(semanticGroup.Visuals, children => semanticGroup.ProjectPaint(children, offsetX, offsetY, paintOrder));
        }
        if (visual is HtmlRenderLayoutRegion layoutRegion) {
            return ProjectGroup(layoutRegion.Visuals, children => layoutRegion.ProjectPaint(children, offsetX, offsetY, paintOrder));
        }
        if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
            return ProjectGroup(logicalTextGroup.Visuals,
                children => logicalTextGroup.ProjectPaint(children, offsetX, offsetY, paintOrder,
                    _ownedLogicalText.Add(logicalTextGroup)));
        }
        if (visual is HtmlRenderFormField formField) {
            return ProjectGroup(formField.Visuals, children => formField.ProjectPaint(children, offsetX, offsetY, paintOrder));
        }
        // Container frames describe flow, not transformed or overflowing child paint.
        // Select leaves in destination coordinates and retain their enclosing graphics state.
        if (visual is HtmlRenderNamedDestination or HtmlRenderBookmarkAnchor) {
            // Navigation has a point and page owner even when its element paints
            // nothing. Ancestor paint clips must not remove an incoming target.
            OfficePoint point = transform.TransformPoint(new OfficePoint(visual.X, visual.Y));
            // Vertical slicing assigns a point to one page, including points
            // outside its horizontal paint bounds. Stitching preserves the page
            // owner already established by layout, even for off-page targets.
            if (!_navigationAlreadyPageOwned && (point.Y < top || point.Y >= bottom)) return null;
            if (!_ownedNavigation.Add(visual)) return null;
        } else if (!Intersects(visual, transform, left, top, right, bottom)) {
            return null;
        }
        if (visual is HtmlRenderText text && !_ownedText.Add(text)) {
            CountVisual();
            HtmlRenderVisual translated = text.TranslatePaint(offsetX, offsetY, 0);
            CountVisual();
            return new HtmlRenderLogicalTextGroup(
                string.Empty,
                translated.X, translated.Y, translated.Width, translated.Height,
                new[] { translated }, paintOrder, translated.Source,
                layoutY: translated.LayoutY, layoutHeight: translated.LayoutHeight);
        }

        CountVisual();
        return visual.TranslatePaint(offsetX, offsetY, paintOrder);

        HtmlRenderVisual? ProjectGroup(
            IEnumerable<HtmlRenderVisual> children,
            Func<IEnumerable<HtmlRenderVisual>, HtmlRenderVisual> create,
            OfficeTransform? childTransform = null) {
            var projectedChildren = new List<HtmlRenderVisual>();
            foreach (HtmlRenderVisual child in children.OrderBy(item => item.PaintOrder)) {
                HtmlRenderVisual? projectedChild = Project(
                    child, left, top, right, bottom, offsetX, offsetY, projectedChildren.Count,
                    childTransform ?? transform);
                if (projectedChild != null) projectedChildren.Add(projectedChild);
            }
            if (projectedChildren.Count == 0) return null;
            CountVisual();
            return create(projectedChildren);
        }

        HtmlRenderVisual? ProjectClipped(IEnumerable<HtmlRenderVisual> children, ProjectionClip clip,
            Func<IEnumerable<HtmlRenderVisual>, HtmlRenderVisual> create) {
            _clips.Add(clip);
            try { return ProjectGroup(children, create); }
            finally { _clips.RemoveAt(_clips.Count - 1); }
        }
    }

    private void CountVisual() {
        _cancellationToken.ThrowIfCancellationRequested();
        if (++_materializedVisuals > _maximumVisuals) {
            throw new InvalidOperationException(
                $"HTML render projection exceeded MaxProjectedVisuals {_maximumVisuals}.");
        }
    }

    private bool Intersects(HtmlRenderVisual visual, OfficeTransform transform,
        double left, double top, double right, double bottom) {
        double overhang = visual is HtmlRenderText text ? text.PaintTopOverflow : 0D;
        var bounds = transform.TransformRectangleBounds(visual.X, visual.Y - overhang, visual.Width, visual.Height + overhang);
        left = Math.Max(left, bounds.Left);
        top = Math.Max(top, bounds.Top);
        right = Math.Min(right, bounds.Right);
        bottom = Math.Min(bottom, bounds.Bottom);
        if (left >= right || top >= bottom) return false;
        if (_clips.Count == 0) return true;

        // Logical ownership requires a common visible region, rather than separate
        // intersections with a slice and each ancestor. Keep that region in destination
        // coordinates while clipping against the ancestors' affine halfplanes.
        var region = new List<OfficePoint> {
            new(left, top), new(right, top), new(right, bottom), new(left, bottom)
        };
        foreach (ProjectionClip clip in _clips) {
            _cancellationToken.ThrowIfCancellationRequested();
            region = clip.Intersect(region);
            if (region.Count < 3) return false;
        }
        double twiceArea = 0D;
        OfficePoint origin = region[0];
        for (int index = 1; index + 1 < region.Count; index++) {
            OfficePoint first = region[index], second = region[index + 1];
            twiceArea += (first.X - origin.X) * (second.Y - origin.Y)
                - (first.Y - origin.Y) * (second.X - origin.X);
        }
        return Math.Abs(twiceArea) > 0.0000000001D;
    }

    private readonly struct ProjectionClip {
        private readonly double _left, _top, _right, _bottom;
        private readonly bool _horizontal, _vertical, _invertible;
        private readonly OfficeTransform _inverse;

        internal ProjectionClip(double x, double y, double width, double height,
            bool horizontal, bool vertical, OfficeTransform transform) {
            _left = x; _top = y; _right = x + width; _bottom = y + height;
            _horizontal = horizontal; _vertical = vertical;
            _invertible = transform.TryInvert(out _inverse);
        }

        internal List<OfficePoint> Intersect(List<OfficePoint> region) {
            if (!_invertible) return new List<OfficePoint>();
            if (_horizontal) {
                region = Clip(region, _inverse.M11, _inverse.M21, _inverse.OffsetX - _left);
                region = Clip(region, -_inverse.M11, -_inverse.M21, _right - _inverse.OffsetX);
            }
            if (_vertical) {
                region = Clip(region, _inverse.M12, _inverse.M22, _inverse.OffsetY - _top);
                region = Clip(region, -_inverse.M12, -_inverse.M22, _bottom - _inverse.OffsetY);
            }
            // Path outlines retain their conservative bounding rectangle here; the
            // existing graphics-state clip performs exact outline clipping during paint.
            return region;
        }

        private static List<OfficePoint> Clip(IReadOnlyList<OfficePoint> input, double a, double b, double c) {
            var output = new List<OfficePoint>(input.Count + 1);
            if (input.Count == 0) return output;
            OfficePoint previous = input[input.Count - 1];
            double previousDistance = a * previous.X + b * previous.Y + c;
            foreach (OfficePoint current in input) {
                double currentDistance = a * current.X + b * current.Y + c;
                if ((previousDistance >= 0D) != (currentDistance >= 0D)) {
                    double fraction = previousDistance / (previousDistance - currentDistance);
                    output.Add(new OfficePoint(previous.X + fraction * (current.X - previous.X),
                        previous.Y + fraction * (current.Y - previous.Y)));
                }
                if (currentDistance >= 0D) output.Add(current);
                previous = current;
                previousDistance = currentDistance;
            }
            return output;
        }
    }
}
