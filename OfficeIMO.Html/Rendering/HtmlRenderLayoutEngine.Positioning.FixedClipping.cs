using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly Dictionary<HtmlRenderVisual, PositionedElementRequest> _fixedClipPaintRequests = new();
    private readonly Dictionary<int, IReadOnlyList<HtmlRenderVisual>> _fixedLegacyClipBoundaries = new();

    private void ResetFixedClipPaintBoundaries() {
        _fixedClipPaintRequests.Clear();
        _fixedLegacyClipBoundaries.Clear();
    }

    private void ApplyFixedAncestorLegacyClips(List<HtmlRenderVisual> visuals) {
        if (_fixedClipPaintRequests.Count == 0) return;
        CaptureFixedLegacyClipBoundaries(visuals, new List<HtmlRenderVisual>());
        for (int index = 0; index < visuals.Count; index++) {
            HtmlRenderVisual visual = visuals[index];
            if (!_fixedClipPaintRequests.TryGetValue(visual, out PositionedElementRequest? request)) continue;
            var ancestors = new Stack<int>();
            for (var ancestor = request.DirectParent; ancestor != null; ancestor = ancestor.ParentElement) {
                if (_semanticNodeIds.TryGetValue(ancestor, out int nodeId)
                    && _fixedLegacyClipBoundaries.ContainsKey(nodeId)) ancestors.Push(nodeId);
            }
            var envelopes = new List<HtmlRenderVisual>();
            var included = new HashSet<HtmlRenderVisual>();
            while (ancestors.Count > 0) {
                foreach (HtmlRenderVisual envelope in _fixedLegacyClipBoundaries[ancestors.Pop()]) {
                    if (included.Add(envelope)) envelopes.Add(envelope);
                }
            }
            if (envelopes.Count == 0) continue;

            // Fixed positioning keeps its existing viewport placement. Express it
            // in the ancestor's paint space so transformed clip edges constrain
            // that placement without applying the ancestor transform twice.
            OfficeTransform transform = OfficeTransform.Identity;
            foreach (HtmlRenderEffectGroup effect in envelopes.OfType<HtmlRenderEffectGroup>()) {
                transform = effect.Transform.Then(transform);
            }
            IReadOnlyList<HtmlRenderVisual> content = ReplaceDescendantFormFieldsForPaintEffect(
                new[] { visual }, "ancestor-clip=" + envelopes.OfType<HtmlRenderClipGroup>().First().Source);
            if (!transform.TryInvert(out OfficeTransform inverse)) {
                visuals[index] = visual.CopyStackingContextTo(new HtmlRenderClipGroup(
                    visual.X, visual.Y, 0D, 0D, true, true, content, visual.PaintOrder, visual.Source)).IdentifyOutOfFlowPaint();
                continue;
            }
            if (inverse != OfficeTransform.Identity) {
                content = new[] { new HtmlRenderEffectGroup(visual.X, visual.Y, visual.Width, visual.Height,
                    inverse, 1D, content, 0, visual.Source) };
            }
            for (int envelopeIndex = envelopes.Count - 1; envelopeIndex >= 0; envelopeIndex--) {
                HtmlRenderVisual envelope = envelopes[envelopeIndex];
                content = new[] { envelope switch {
                    HtmlRenderClipGroup clip => clip.ProjectPaint(content, 0D, 0D, 0),
                    HtmlRenderPathClipGroup path => path.ProjectPaint(content, 0D, 0D, 0),
                    HtmlRenderEffectGroup effect => effect.ProjectPaint(content, 0D, 0D, 0),
                    _ => throw new InvalidOperationException("Unsupported fixed clipping boundary.")
                } };
            }
            visuals[index] = visual.CopyStackingContextTo(content[0].Translate(0D, 0D, visual.PaintOrder)).IdentifyOutOfFlowPaint();
        }
        _fixedClipPaintRequests.Clear();
        // Absolute ancestors paint on the first print page, while fixed children
        // repeat. Retain their exact clip boundaries for the later pages too.
    }

    private void CaptureFixedLegacyClipBoundaries(IEnumerable<HtmlRenderVisual> visuals, List<HtmlRenderVisual> path) {
        foreach (HtmlRenderVisual visual in visuals) {
            CheckCancellation();
            bool envelope = visual is HtmlRenderEffectGroup or HtmlRenderPathClipGroup
                || visual is HtmlRenderClipGroup { LegacyClipOwnerNodeId: not null };
            if (envelope) path.Add(visual);
            if (visual is HtmlRenderClipGroup { LegacyClipOwnerNodeId: int nodeId } clip) {
                var boundary = new List<HtmlRenderVisual>(path);
                // clip-path is inside legacy clip on the same authored box.
                if (clip.Visuals.Count == 1 && clip.Visuals[0] is HtmlRenderPathClipGroup sameBox
                    && sameBox.Source == clip.Source) boundary.Add(sameBox);
                _fixedLegacyClipBoundaries[nodeId] = boundary;
            }
            IReadOnlyList<HtmlRenderVisual>? children = GetGroupChildren(visual);
            if (children != null) CaptureFixedLegacyClipBoundaries(children, path);
            if (envelope) path.RemoveAt(path.Count - 1);
        }
    }
}
