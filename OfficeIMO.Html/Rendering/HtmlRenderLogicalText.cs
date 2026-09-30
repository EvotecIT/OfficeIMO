using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Html;

internal static class HtmlRenderLogicalText {
    internal static bool ContainsText(HtmlRenderVisual visual) {
        var fragments = new List<LogicalTextFragment>();
        CollectLogicalTextFragments(new[] { visual }, fragments);
        return fragments.Any(fragment => fragment.Text.Length > 0);
    }

    internal static bool TryResolveSourceText(IEnumerable<HtmlRenderVisual> visuals, out string logicalText, bool preserveBlockSeparators = false) {
        var fragments = new List<LogicalTextFragment>();
        CollectLogicalTextFragments(visuals, fragments);
        logicalText = string.Empty;
        if (fragments.Count == 0 || fragments.Any(fragment => !fragment.Order.HasValue)) return false;
        var text = new System.Text.StringBuilder();
        object? previousBlock = null;
        foreach (LogicalTextFragment fragment in fragments.OrderBy(fragment => fragment.Order!.Value).ThenBy(fragment => fragment.PaintSequence)) {
            if (preserveBlockSeparators && text.Length > 0 && !Equals(previousBlock, fragment.BlockIdentity)
                && (previousBlock != null || fragment.BlockIdentity != null)) text.Append('\n');
            text.Append(fragment.Text);
            previousBlock = fragment.BlockIdentity;
        }
        logicalText = text.ToString();
        return logicalText.Length > 0;
    }

    internal static bool TryResolveSourceOrderRange(HtmlRenderVisual visual, out int first, out int last) {
        var fragments = new List<LogicalTextFragment>();
        CollectLogicalTextFragments(new[] { visual }, fragments);
        fragments.RemoveAll(fragment => fragment.Text.Length == 0);
        first = last = 0;
        if (fragments.Count == 0 || fragments.Any(fragment => !fragment.Order.HasValue)) return false;
        first = fragments.Min(fragment => fragment.Order!.Value);
        last = fragments.Max(fragment => fragment.Order!.Value);
        return true;
    }

    internal static bool TryResolveReorderedText(IEnumerable<HtmlRenderVisual> visuals, out string logicalText) {
        var fragments = new List<LogicalTextFragment>();
        CollectLogicalTextFragments(visuals, fragments);
        logicalText = string.Empty;
        if (fragments.Count < 2 || fragments.Any(fragment => !fragment.Order.HasValue)) return false;

        List<LogicalTextFragment> ordered = fragments
            .OrderBy(fragment => fragment.Order!.Value)
            .ThenBy(fragment => fragment.PaintSequence)
            .ToList();
        if (ordered.Select(fragment => fragment.PaintSequence).SequenceEqual(fragments.Select(fragment => fragment.PaintSequence))) return false;
        logicalText = string.Concat(ordered.Select(fragment => fragment.Text));
        return logicalText.Length > 0;
    }

    private static void CollectLogicalTextFragments(IEnumerable<HtmlRenderVisual> visuals, ICollection<LogicalTextFragment> fragments, object? blockIdentity = null) {
        foreach (HtmlRenderVisual visual in visuals.OrderBy(item => item.PaintOrder)) {
            if (visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Artifact }) continue;
            if (visual is HtmlRenderLogicalTextGroup logicalTextGroup) {
                if (logicalTextGroup.IsFlowAnchor) continue;
                if (ContainsArtifactVisual(logicalTextGroup.Visuals)) {
                    CollectLogicalTextFragments(logicalTextGroup.Visuals, fragments, blockIdentity);
                    continue;
                }
                int? order = ResolveLogicalTextOrder(logicalTextGroup.Visuals);
                fragments.Add(new LogicalTextFragment(logicalTextGroup.Text, order, fragments.Count, blockIdentity));
                continue;
            }
            if (visual is HtmlRenderText text) {
                fragments.Add(new LogicalTextFragment(text.Text, text.LogicalTextOrder, fragments.Count, blockIdentity));
                continue;
            }

            IEnumerable<HtmlRenderVisual>? children = LogicalTextChildVisuals(visual);
            if (children != null) CollectLogicalTextFragments(children, fragments,
                visual is HtmlRenderSemanticGroup semantic && HtmlRenderSemanticGroup.IsTextContentRole(semantic.Role)
                    ? (object?)semantic.StructureElementKey ?? semantic : blockIdentity);
        }
    }

    private static bool ContainsArtifactVisual(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Artifact }) return true;
            IEnumerable<HtmlRenderVisual>? children = LogicalTextChildVisuals(visual);
            if (children != null && ContainsArtifactVisual(children)) return true;
        }
        return false;
    }

    private static int? ResolveLogicalTextOrder(IEnumerable<HtmlRenderVisual> visuals) {
        int? order = null;
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderSemanticGroup { Role: HtmlRenderSemanticGroupRole.Artifact }) continue;
            int? candidate = visual is HtmlRenderText text
                ? text.LogicalTextOrder
                : LogicalTextChildVisuals(visual) is IEnumerable<HtmlRenderVisual> children
                    ? ResolveLogicalTextOrder(children)
                    : null;
            if (candidate.HasValue && (!order.HasValue || candidate.Value < order.Value)) order = candidate;
        }
        return order;
    }

    private static IEnumerable<HtmlRenderVisual>? LogicalTextChildVisuals(HtmlRenderVisual visual) => visual is HtmlRenderClipGroup clipGroup
        ? clipGroup.Visuals
        : visual is HtmlRenderLayoutRegion layoutRegion
            ? layoutRegion.Visuals
        : visual is HtmlRenderPathClipGroup pathClipGroup
            ? pathClipGroup.Visuals
            : visual is HtmlRenderEffectGroup effectGroup
                ? effectGroup.Visuals
                : visual is HtmlRenderSemanticGroup semanticGroup
                    ? semanticGroup.Visuals
                    : visual is HtmlRenderLogicalTextGroup logicalTextGroup
                        ? logicalTextGroup.Visuals
                        : visual is HtmlRenderFormField formField ? formField.Visuals : null;

    private readonly struct LogicalTextFragment {
        internal LogicalTextFragment(string text, int? order, int paintSequence, object? blockIdentity) {
            Text = text;
            Order = order;
            PaintSequence = paintSequence;
            BlockIdentity = blockIdentity;
        }

        internal string Text { get; }
        internal int? Order { get; }
        internal int PaintSequence { get; }
        internal object? BlockIdentity { get; }
    }
}
