using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Html.Pdf;

/// <summary>Operation-local link coverage and logical ownership for one rendered page.</summary>
internal sealed class HtmlPdfPagePaintContext {
    private readonly HtmlPdfAnchorLinkMap _links;
    internal HtmlPdfMathMlFiles MathMlFiles { get; }
    private readonly Dictionary<HtmlRenderLogicalTextScope, List<HtmlRenderLogicalTextGroup>> _logicalGroups = new();
    private readonly HashSet<HtmlRenderLogicalTextScope> _emittedScopes = new();
    private readonly Dictionary<HtmlRenderLogicalTextScope, IReadOnlyList<HtmlRenderVisual>> _logicalPaint = new();

    private HtmlPdfPagePaintContext(HtmlRenderPage page, HtmlPdfMathMlFiles mathMlFiles) {
        MathMlFiles = mathMlFiles;
        _links = HtmlPdfAnchorLinkMap.Create(page);
        Collect(page.Scene);
    }

    internal static HtmlPdfPagePaintContext Create(HtmlRenderPage page, HtmlPdfMathMlFiles mathMlFiles) => new(page, mathMlFiles);

    internal bool IsActive(HtmlRenderAnchorFragment fragment) => _links.IsActive(fragment);
    internal bool Covers(HtmlRenderVisual visual) => _links.Covers(visual);
    internal bool TryClaim(HtmlRenderLogicalTextScope scope) => _emittedScopes.Add(scope);
    internal bool IsClaimed(HtmlRenderLogicalTextScope scope) => _emittedScopes.Contains(scope);

    internal IReadOnlyList<HtmlRenderVisual> GetLogicalPaint(HtmlRenderLogicalTextScope scope) {
        if (!_logicalPaint.TryGetValue(scope, out IReadOnlyList<HtmlRenderVisual>? paint)) {
            paint = _logicalGroups[scope].SelectMany(group => group.Visuals).ToArray();
            _logicalPaint.Add(scope, paint);
        }
        return paint;
    }

    private void Collect(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderLogicalTextGroup logical && logical.LogicalScope is { } scope) {
                if (!_logicalGroups.TryGetValue(scope, out List<HtmlRenderLogicalTextGroup>? groups)) {
                    groups = new List<HtmlRenderLogicalTextGroup>();
                    _logicalGroups.Add(scope, groups);
                }
                groups.Add(logical);
            }
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLogicalTextGroup text => text.Visuals,
                HtmlRenderLayoutRegion region => region.Visuals,
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup clip => clip.Visuals,
                HtmlRenderEffectGroup effect => effect.Visuals,
                HtmlRenderFormField field => field.Visuals,
                _ => null
            };
            if (children != null) Collect(children);
        }
    }
}
