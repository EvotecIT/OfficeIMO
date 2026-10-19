using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html.Pdf;

/// <summary>Finds paint links covered by anchor-owned fragments without conflating equal URLs.</summary>
internal sealed class HtmlPdfAnchorLinkMap {
    private readonly HashSet<HtmlRenderVisual> _coveredVisuals;
    private readonly HashSet<HtmlRenderAnchorFragment> _activeFragments;

    private HtmlPdfAnchorLinkMap(HashSet<HtmlRenderVisual> coveredVisuals,
        HashSet<HtmlRenderAnchorFragment> activeFragments) {
        _coveredVisuals = coveredVisuals;
        _activeFragments = activeFragments;
    }

    internal static HtmlPdfAnchorLinkMap Create(HtmlRenderPage page) {
        var grouped = new Dictionary<(object Space, string Uri), List<HtmlRenderRectangle>>();
        var linkedVisuals = new Dictionary<HtmlRenderVisual, (object Space, string Uri, HtmlRenderRectangle Bounds)>();
        var activeFragments = new HashSet<HtmlRenderAnchorFragment>();
        Collect(page.Scene, new object(), RectangularClip.Unbounded, grouped, linkedVisuals, activeFragments);
        var regions = grouped.ToDictionary(
            pair => pair.Key,
            pair => new LinkRegionSet(pair.Value));
        var covered = new HashSet<HtmlRenderVisual>();
        foreach (var entry in linkedVisuals) {
            HtmlRenderVisual visual = entry.Key;
            if (regions.TryGetValue((entry.Value.Space, entry.Value.Uri), out LinkRegionSet? set)
                && set.Contains(entry.Value.Bounds)) covered.Add(visual);
        }
        return new HtmlPdfAnchorLinkMap(covered, activeFragments);
    }

    internal bool IsActive(HtmlRenderAnchorFragment fragment) => _activeFragments.Contains(fragment);

    internal bool Covers(HtmlRenderVisual visual) => _coveredVisuals.Contains(visual);

    private static void Collect(IEnumerable<HtmlRenderVisual> visuals,
        object space,
        RectangularClip clipBounds,
        IDictionary<(object Space, string Uri), List<HtmlRenderRectangle>> grouped,
        IDictionary<HtmlRenderVisual, (object Space, string Uri, HtmlRenderRectangle Bounds)> linkedVisuals,
        ISet<HtmlRenderAnchorFragment> activeFragments) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderAnchorFragment fragment) {
                activeFragments.Add(fragment);
                HtmlRenderRectangle? bounds = clipBounds.Intersect(new HtmlRenderRectangle(fragment.X, fragment.Y, fragment.Width, fragment.Height));
                if (!bounds.HasValue) continue;
                var key = (space, fragment.LinkUri!);
                if (!grouped.TryGetValue(key, out List<HtmlRenderRectangle>? regions)) {
                    regions = new List<HtmlRenderRectangle>();
                    grouped.Add(key, regions);
                }
                regions.Add(bounds.Value);
            } else if (visual is HtmlRenderSemanticGroup semantic) Collect(semantic.Visuals, space, clipBounds, grouped, linkedVisuals, activeFragments);
            else if (visual is HtmlRenderLogicalTextGroup logical) Collect(logical.Visuals, space, clipBounds, grouped, linkedVisuals, activeFragments);
            else if (visual is HtmlRenderLayoutRegion layout) Collect(layout.Visuals, space, clipBounds, grouped, linkedVisuals, activeFragments);
            else if (visual is HtmlRenderClipGroup clip) Collect(clip.Visuals, space, clipBounds.Constrain(clip.ClipHorizontal ? clip.ClipX : double.NegativeInfinity,
                clip.ClipVertical ? clip.ClipY : double.NegativeInfinity,
                clip.ClipHorizontal ? clip.ClipX + clip.ClipWidth : double.PositiveInfinity,
                clip.ClipVertical ? clip.ClipY + clip.ClipHeight : double.PositiveInfinity), grouped, linkedVisuals, activeFragments);
            // Rectangular path clips retain the exact anchor rectangle intersection
            // in the PDF canvas, including nested overflow and legacy clip rectangles.
            else if (visual is HtmlRenderPathClipGroup rectangular && rectangular.ClipPath.Kind == OfficeClipPathKind.Rectangle)
                Collect(rectangular.Visuals, space, clipBounds.Constrain(rectangular.ClipX, rectangular.ClipY,
                    rectangular.ClipX + rectangular.ClipPath.Width, rectangular.ClipY + rectangular.ClipPath.Height), grouped, linkedVisuals, activeFragments);
            // PDF annotation rectangles do not inherit a nonrectangular path clip.
            // Keep the existing per-visual link handling inside that path instead.
            else if (visual is HtmlRenderPathClipGroup) continue;
            // Child rectangles are local to this effect. An equal URI in another
            // transform space does not establish annotation coverage here.
            else if (visual is HtmlRenderEffectGroup effect) Collect(effect.Visuals, effect, RectangularClip.Unbounded, grouped, linkedVisuals, activeFragments);
            else if (visual is HtmlRenderFormField field) Collect(field.Visuals, space, clipBounds, grouped, linkedVisuals, activeFragments);
            else if (visual.LinkUri != null) {
                HtmlRenderRectangle? bounds = clipBounds.Intersect(visual is HtmlRenderText text && text.LinkBounds.HasValue
                    ? text.LinkBounds.Value : new HtmlRenderRectangle(visual.X, visual.Y, visual.Width, visual.Height));
                if (bounds.HasValue) linkedVisuals[visual] = (space, visual.LinkUri, bounds.Value);
            }
        }
    }

    // Coverage describes rectangles that the canvas can actually annotate, rather
    // than an unclipped source box that may belong to a different equal-URI anchor.
    private readonly struct RectangularClip {
        private readonly double _left, _top, _right, _bottom;
        private RectangularClip(double left, double top, double right, double bottom) {
            _left = left; _top = top; _right = right; _bottom = bottom;
        }
        internal static RectangularClip Unbounded => new(double.NegativeInfinity, double.NegativeInfinity,
            double.PositiveInfinity, double.PositiveInfinity);
        internal RectangularClip Constrain(double left, double top, double right, double bottom) =>
            new(Math.Max(_left, left), Math.Max(_top, top), Math.Min(_right, right), Math.Min(_bottom, bottom));
        internal HtmlRenderRectangle? Intersect(HtmlRenderRectangle rectangle) {
            double left = Math.Max(_left, rectangle.X), top = Math.Max(_top, rectangle.Y);
            double right = Math.Min(_right, rectangle.Right), bottom = Math.Min(_bottom, rectangle.Bottom);
            return right > left && bottom > top ? new HtmlRenderRectangle(left, top, right - left, bottom - top) : null;
        }
    }

    private sealed class LinkRegionSet {
        private readonly HtmlRenderRectangle[] _fragments;
        private readonly double[] _maximumBottom;

        internal LinkRegionSet(IEnumerable<HtmlRenderRectangle> fragments) {
            _fragments = fragments.OrderBy(fragment => fragment.Y).ToArray();
            _maximumBottom = new double[_fragments.Length];
            double maximum = double.NegativeInfinity;
            for (int index = 0; index < _fragments.Length; index++) {
                maximum = Math.Max(maximum, _fragments[index].Y + _fragments[index].Height);
                _maximumBottom[index] = maximum;
            }
        }

        internal bool Contains(HtmlRenderRectangle visual) {
            // Only accommodate floating-point arithmetic; authored fractional overflow remains linked.
            const double edgeTolerance = 0.0000001D;
            double top = visual.Y - edgeTolerance;
            double bottom = visual.Y + visual.Height + edgeTolerance;
            int low = 0;
            int high = _fragments.Length;
            while (low < high) {
                int middle = low + (high - low) / 2;
                if (_maximumBottom[middle] < top) low = middle + 1;
                else high = middle;
            }
            for (int index = low; index < _fragments.Length && _fragments[index].Y <= bottom; index++) {
                HtmlRenderRectangle fragment = _fragments[index];
                if (fragment.Y <= visual.Y + edgeTolerance
                    && fragment.Y + fragment.Height >= visual.Y + visual.Height - edgeTolerance
                    && fragment.X <= visual.X + edgeTolerance
                    && fragment.X + fragment.Width >= visual.X + visual.Width - edgeTolerance) return true;
            }
            return false;
        }
    }
}
