namespace OfficeIMO.Html.Pdf;

/// <summary>Operation-local link coverage and logical ownership for one rendered page.</summary>
internal sealed class HtmlPdfPagePaintContext {
    private readonly HtmlPdfAnchorLinkMap _links;
    internal HtmlPdfRenderedConverter.LinkTargets LinkTargets { get; }
    internal HtmlPdfMathMlFiles MathMlFiles { get; }
    private readonly Dictionary<HtmlRenderLogicalTextScope, List<HtmlRenderLogicalTextGroup>> _logicalGroups = new();
    private readonly HashSet<HtmlRenderLogicalTextScope> _emittedScopes = new();
    private readonly Dictionary<HtmlRenderLogicalTextScope, IReadOnlyList<HtmlRenderVisual>> _logicalPaint = new();

    private HtmlPdfPagePaintContext(HtmlRenderPage page, HtmlPdfMathMlFiles mathMlFiles, HtmlPdfRenderedConverter.LinkTargets linkTargets) {
        MathMlFiles = mathMlFiles;
        LinkTargets = linkTargets;
        _links = HtmlPdfAnchorLinkMap.Create(page);
        Collect(page.Scene);
    }

    internal static HtmlPdfPagePaintContext Create(HtmlRenderPage page, HtmlPdfMathMlFiles mathMlFiles,
        HtmlPdfRenderedConverter.LinkTargets linkTargets) => new(page, mathMlFiles, linkTargets);

    internal bool IsActive(HtmlRenderAnchorFragment fragment) => _links.IsActive(fragment)
        && LinkTargets.Include(fragment.LinkUri, fragment.Source);
    internal bool SuppressesLink(HtmlRenderVisual visual) => _links.Covers(visual)
        || !LinkTargets.Include(visual.LinkUri, visual.Source);
    internal bool TryClaim(HtmlRenderLogicalTextScope scope) => _emittedScopes.Add(scope);
    internal bool IsClaimed(HtmlRenderLogicalTextScope scope) => _emittedScopes.Contains(scope);

    internal IReadOnlyList<HtmlRenderVisual> GetLogicalPaint(HtmlRenderLogicalTextScope scope) {
        if (!_logicalPaint.TryGetValue(scope, out IReadOnlyList<HtmlRenderVisual>? paint)) {
            paint = _logicalGroups[scope].SelectMany(group => group.Visuals).ToArray();
            _logicalPaint.Add(scope, paint);
        }
        return paint;
    }

    private void Collect(IReadOnlyList<HtmlRenderVisual> visuals) {
        for (int index = 0; index < visuals.Count; index++) {
            HtmlRenderVisual visual = visuals[index];
            if (visual is HtmlRenderLogicalTextGroup logical && logical.LogicalScope is { } scope) {
                if (!_logicalGroups.TryGetValue(scope, out List<HtmlRenderLogicalTextGroup>? groups)) {
                    groups = new List<HtmlRenderLogicalTextGroup>();
                    _logicalGroups.Add(scope, groups);
                }
                groups.Add(logical);
            }
            IReadOnlyList<HtmlRenderVisual>? children = visual switch {
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
