using System.Collections.Generic;
using System.Linq;

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
        var grouped = new Dictionary<(object Space, string Uri), List<HtmlRenderAnchorFragment>>();
        var linkedVisuals = new Dictionary<HtmlRenderVisual, (object Space, string Uri)>();
        var activeFragments = new HashSet<HtmlRenderAnchorFragment>();
        Collect(page.Scene, new object(), grouped, linkedVisuals, activeFragments);
        var regions = grouped.ToDictionary(
            pair => pair.Key,
            pair => new LinkRegionSet(pair.Value));
        var covered = new HashSet<HtmlRenderVisual>();
        foreach (var entry in linkedVisuals) {
            HtmlRenderVisual visual = entry.Key;
            if (regions.TryGetValue(entry.Value, out LinkRegionSet? set)
                && set.Contains(visual is HtmlRenderText text && text.LinkBounds.HasValue
                    ? text.LinkBounds.Value
                    : new HtmlRenderRectangle(visual.X, visual.Y, visual.Width, visual.Height))) covered.Add(visual);
        }
        return new HtmlPdfAnchorLinkMap(covered, activeFragments);
    }

    internal bool IsActive(HtmlRenderAnchorFragment fragment) => _activeFragments.Contains(fragment);

    internal bool Covers(HtmlRenderVisual visual) => _coveredVisuals.Contains(visual);

    private static void Collect(IEnumerable<HtmlRenderVisual> visuals,
        object space,
        IDictionary<(object Space, string Uri), List<HtmlRenderAnchorFragment>> grouped,
        IDictionary<HtmlRenderVisual, (object Space, string Uri)> linkedVisuals,
        ISet<HtmlRenderAnchorFragment> activeFragments) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderAnchorFragment fragment) {
                var key = (space, fragment.LinkUri!);
                if (!grouped.TryGetValue(key, out List<HtmlRenderAnchorFragment>? regions)) {
                    regions = new List<HtmlRenderAnchorFragment>();
                    grouped.Add(key, regions);
                }
                regions.Add(fragment);
                activeFragments.Add(fragment);
            } else if (visual is HtmlRenderSemanticGroup semantic) Collect(semantic.Visuals, space, grouped, linkedVisuals, activeFragments);
            else if (visual is HtmlRenderLogicalTextGroup logical) Collect(logical.Visuals, space, grouped, linkedVisuals, activeFragments);
            else if (visual is HtmlRenderLayoutRegion layout) Collect(layout.Visuals, space, grouped, linkedVisuals, activeFragments);
            else if (visual is HtmlRenderClipGroup clip) Collect(clip.Visuals, space, grouped, linkedVisuals, activeFragments);
            // PDF annotation rectangles do not inherit a nonrectangular path clip.
            // Keep the existing per-visual link handling inside that path instead.
            else if (visual is HtmlRenderPathClipGroup) continue;
            // Child rectangles are local to this effect. An equal URI in another
            // transform space does not establish annotation coverage here.
            else if (visual is HtmlRenderEffectGroup effect) Collect(effect.Visuals, effect, grouped, linkedVisuals, activeFragments);
            else if (visual is HtmlRenderFormField field) Collect(field.Visuals, space, grouped, linkedVisuals, activeFragments);
            else if (visual.LinkUri != null) linkedVisuals[visual] = (space, visual.LinkUri);
        }
    }

    private sealed class LinkRegionSet {
        private readonly HtmlRenderAnchorFragment[] _fragments;
        private readonly double[] _maximumBottom;

        internal LinkRegionSet(IEnumerable<HtmlRenderAnchorFragment> fragments) {
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
                HtmlRenderAnchorFragment fragment = _fragments[index];
                if (fragment.Y <= visual.Y + edgeTolerance
                    && fragment.Y + fragment.Height >= visual.Y + visual.Height - edgeTolerance
                    && fragment.X <= visual.X + edgeTolerance
                    && fragment.X + fragment.Width >= visual.X + visual.Width - edgeTolerance) return true;
            }
            return false;
        }
    }
}
