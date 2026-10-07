namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Reuses split-paint text ownership when an extracted float interrupts a source paragraph or container.</summary>
    private void PreservePageFloatLogicalOwnership(List<HtmlRenderVisual> visuals, HashSet<int> floatIndices) {
        if (floatIndices.Count == 0) return;
        var ranges = new Dictionary<int, (int First, int Last)>();
        for (int index = 0; index < visuals.Count; index++) {
            ChargeLayoutOperation("page-float logical text");
            if (HtmlRenderLogicalText.TryResolveSourceOrderRange(visuals[index], out int first, out int last)) ranges.Add(index, (first, last));
        }
        var scopes = new List<HashSet<int>>();
        foreach (int floatIndex in floatIndices) {
            if (!ranges.TryGetValue(floatIndex, out var floated)) continue;
            var members = new HashSet<int> { floatIndex };
            foreach (var pair in ranges) {
                ChargeLayoutOperation("page-float logical ownership");
                if (!floatIndices.Contains(pair.Key) && pair.Value.First < floated.First && pair.Value.Last > floated.Last) members.Add(pair.Key);
            }
            if (members.Count == 1) continue;
            for (int index = 0; index < scopes.Count;) {
                ChargeLayoutOperation("page-float logical scopes");
                if (!scopes[index].Overlaps(members)) {
                    index++;
                    continue;
                }
                members.UnionWith(scopes[index]);
                scopes.RemoveAt(index);
                index = 0;
            }
            scopes.Add(members);
        }
        foreach (HashSet<int> members in scopes) {
            var sourcePaint = members.OrderBy(index => visuals[index].PaintOrder).Select(index => visuals[index]).ToArray();
            if (!HtmlRenderLogicalText.TryResolveSourceText(sourcePaint, out string text, preserveBlockSeparators: true)) continue;
            int textOwner = members.OrderBy(index => ranges[index].First).ThenBy(index => index).First();
            var scope = new HtmlRenderLogicalTextScope(preserveBlockSeparators: true);
            foreach (int index in members) {
                HtmlRenderVisual visual = visuals[index];
                visuals[index] = visual.CopyStackingContextTo(new HtmlRenderLogicalTextGroup(
                    index == textOwner ? text : string.Empty, visual.X, visual.Y, visual.Width, visual.Height,
                    new[] { visual }, visual.PaintOrder, visual.Source, visual.LayoutY, visual.LayoutHeight, scope));
            }
        }
    }
}
