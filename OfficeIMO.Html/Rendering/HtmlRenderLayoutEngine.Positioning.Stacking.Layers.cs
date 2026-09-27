namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static IEnumerable<HtmlRenderVisual> EnumerateRootStackingLayers(HtmlRenderVisual visual) {
        // A context is atomic: its descendants never compete with root siblings.
        if (visual.StackingContext != null) {
            yield return visual;
            yield break;
        }
        IReadOnlyList<HtmlRenderVisual>? children = visual switch {
            HtmlRenderClipGroup clip => clip.Visuals,
            HtmlRenderPathClipGroup clip => clip.Visuals,
            HtmlRenderSemanticGroup semantic => semantic.Visuals,
            HtmlRenderLayoutRegion region => region.Visuals,
            HtmlRenderLogicalTextGroup text => text.Visuals,
            _ => null
        };
        if (children == null || !children.Any(HasNestedStackingContext)) {
            yield return visual;
            yield break;
        }
        List<HtmlRenderVisual> layers = children.SelectMany(EnumerateRootStackingLayers).ToList();
        string reorderedText = string.Empty;
        bool hasReorderedText = visual is HtmlRenderSemanticGroup semanticText
            && HtmlRenderSemanticGroup.IsTextContentRole(semanticText.Role)
            && HtmlRenderLogicalText.TryResolveReorderedText(layers
                .OrderBy(layer => layer.StackingContext == null ? 1 : layer.StackingContext.ZIndex < 0 ? 0 : 2)
                .ThenBy(layer => layer.StackingContext?.ZIndex ?? 0)
                .ThenBy(layer => layer.StackingContext?.SourceOrder ?? -1)
                .Select((layer, index) => layer.Translate(0D, 0D, index)), out reorderedText);
        HtmlRenderLogicalTextScope? logicalScope = hasReorderedText ? new HtmlRenderLogicalTextScope() : null;
        bool ownsLogicalText = true;
        bool reorderedTextAssigned = false;
        foreach (HtmlRenderVisual layer in layers) {
            HtmlRenderVisual projected = layer.Translate(0D, 0D, 0);
            if (hasReorderedText && HtmlRenderLogicalText.ContainsText(layer)) {
                // Paint fragments retain one source-order replacement across the
                // split. Later pieces are paint-only, as with bidi glyph fragments.
                projected = new HtmlRenderLogicalTextGroup(!reorderedTextAssigned ? reorderedText : string.Empty,
                    visual.X, visual.Y, visual.Width, visual.Height, new[] { projected },
                    0, visual.Source, visual.LayoutY, visual.LayoutHeight, logicalScope);
                reorderedTextAssigned = true;
            }
            var content = new[] { projected };
            HtmlRenderVisual envelope = visual switch {
                HtmlRenderClipGroup clip => clip.ProjectPaint(content, 0D, 0D, visual.PaintOrder),
                HtmlRenderPathClipGroup clip => clip.ProjectPaint(content, 0D, 0D, visual.PaintOrder),
                HtmlRenderSemanticGroup semantic => semantic.ProjectPaint(content, 0D, 0D, visual.PaintOrder),
                HtmlRenderLayoutRegion region => region.ProjectPaint(content, 0D, 0D, visual.PaintOrder),
                HtmlRenderLogicalTextGroup text => text.ProjectPaint(content, 0D, 0D, visual.PaintOrder, ownsLogicalText),
                _ => throw new InvalidOperationException("An unsupported root paint envelope was split.")
            };
            // Keep each ancestor's clip and semantic identity around its promoted
            // layer; only the external paint ordering changes.
            yield return layer.CopyStackingContextTo(envelope);
            ownsLogicalText = false;
        }
    }

    private static bool HasNestedStackingContext(HtmlRenderVisual visual) {
        if (visual.StackingContext != null) return true;
        IEnumerable<HtmlRenderVisual>? children = visual switch {
            HtmlRenderClipGroup clip => clip.Visuals,
            HtmlRenderPathClipGroup clip => clip.Visuals,
            HtmlRenderSemanticGroup semantic => semantic.Visuals,
            HtmlRenderLayoutRegion region => region.Visuals,
            HtmlRenderLogicalTextGroup text => text.Visuals,
            _ => null
        };
        return children?.Any(HasNestedStackingContext) == true;
    }
}
