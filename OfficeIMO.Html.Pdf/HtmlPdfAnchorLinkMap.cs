using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Html.Pdf;

/// <summary>Finds paint links covered by anchor-owned fragments without conflating equal URLs.</summary>
internal sealed class HtmlPdfAnchorLinkMap {
    private readonly Dictionary<string, LinkRegionSet> _byUri;
    private readonly HashSet<HtmlRenderAnchorFragment> _activeFragments;

    private HtmlPdfAnchorLinkMap(Dictionary<string, LinkRegionSet> byUri,
        HashSet<HtmlRenderAnchorFragment> activeFragments) {
        _byUri = byUri;
        _activeFragments = activeFragments;
    }

    internal static HtmlPdfAnchorLinkMap Create(HtmlRenderPage page) {
        var grouped = new Dictionary<string, List<HtmlRenderAnchorFragment>>(StringComparer.Ordinal);
        var activeFragments = new HashSet<HtmlRenderAnchorFragment>();
        Collect(page.Scene, grouped, activeFragments);
        return new HtmlPdfAnchorLinkMap(grouped.ToDictionary(
            pair => pair.Key,
            pair => new LinkRegionSet(pair.Value),
            StringComparer.Ordinal), activeFragments);
    }

    internal bool IsActive(HtmlRenderAnchorFragment fragment) => _activeFragments.Contains(fragment);

    internal bool Covers(HtmlRenderVisual visual) =>
        visual is not HtmlRenderAnchorFragment
        && visual.LinkUri != null
        && _byUri.TryGetValue(visual.LinkUri, out LinkRegionSet? regions)
        && regions.Overlaps(visual);

    private static void Collect(IEnumerable<HtmlRenderVisual> visuals,
        IDictionary<string, List<HtmlRenderAnchorFragment>> grouped,
        ISet<HtmlRenderAnchorFragment> activeFragments) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderAnchorFragment fragment) {
                if (!grouped.TryGetValue(fragment.LinkUri!, out List<HtmlRenderAnchorFragment>? regions)) {
                    regions = new List<HtmlRenderAnchorFragment>();
                    grouped.Add(fragment.LinkUri!, regions);
                }
                regions.Add(fragment);
                activeFragments.Add(fragment);
            } else if (visual is HtmlRenderSemanticGroup semantic) Collect(semantic.Visuals, grouped, activeFragments);
            else if (visual is HtmlRenderLogicalTextGroup logical) Collect(logical.Visuals, grouped, activeFragments);
            else if (visual is HtmlRenderLayoutRegion layout) Collect(layout.Visuals, grouped, activeFragments);
            else if (visual is HtmlRenderClipGroup clip) Collect(clip.Visuals, grouped, activeFragments);
            // PDF annotation rectangles do not inherit a nonrectangular path clip.
            // Keep the existing per-visual link handling inside that path instead.
            else if (visual is HtmlRenderPathClipGroup) continue;
            else if (visual is HtmlRenderEffectGroup effect) Collect(effect.Visuals, grouped, activeFragments);
            else if (visual is HtmlRenderFormField field) Collect(field.Visuals, grouped, activeFragments);
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

        internal bool Overlaps(HtmlRenderVisual visual) {
            const double edgeTolerance = 1D;
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
                if (fragment.X <= visual.X + visual.Width + edgeTolerance
                    && fragment.X + fragment.Width >= visual.X - edgeTolerance) return true;
            }
            return false;
        }
    }
}
