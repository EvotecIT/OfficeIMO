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
            visuals[index] = ApplyFixedLegacyClipEnvelopes(visual,
                ResolveFixedClipEnvelopes(request, _fixedLegacyClipBoundaries)).IdentifyOutOfFlowPaint();
        }
        _fixedClipPaintRequests.Clear();
        // Absolute ancestors paint on the first print page, while fixed children
        // repeat. Retain their exact clip boundaries for the later pages too.
    }

    private IReadOnlyList<HtmlRenderVisual> ResolveLocalFixedAncestorClips(
        PositionedElementRequest request, double width, double height, double originX, double originY) {
        if (request.Style.Position != "fixed" || request.IsViewportFixed
            || ReferenceEquals(request.DirectParent, request.ContainingBlock)) return Array.Empty<HtmlRenderVisual>();
        var boundaries = new Dictionary<int, IReadOnlyList<HtmlRenderVisual>>();
        if (_localPositionedElements.TryGetValue(request.ContainingBlock, out List<PositionedElementRequest>? requests)) {
            foreach (PositionedElementRequest ancestor in requests.Where(candidate =>
                         !ReferenceEquals(candidate, request) && ContainsElementOrSelf(candidate.Element, request.DirectParent))) {
                bool hasRect = _positionedContainingRects.TryGetValue(ancestor.Element, out PositionedContainingRect? rect);
                PositionedLayer layer = ancestor.Resolve(this, hasRect ? rect!.Width : width, hasRect ? rect!.Height : height);
                CaptureFixedLegacyClipBoundaries(layer.Block.Visuals.Select(visual => visual.Translate(
                    originX + (hasRect ? rect!.X : 0D) + layer.X,
                    originY + (hasRect ? rect!.Y : 0D) + layer.Y, visual.PaintOrder)), new List<HtmlRenderVisual>(), boundaries);
            }
        }
        return ResolveFixedClipEnvelopes(request, boundaries, stopAtContainingBlock: true);
    }

    private IReadOnlyList<HtmlRenderVisual> ResolveFixedClipEnvelopes(PositionedElementRequest request,
        IReadOnlyDictionary<int, IReadOnlyList<HtmlRenderVisual>> boundaries, bool stopAtContainingBlock = false) {
        var ancestors = new Stack<int>();
        for (var ancestor = request.DirectParent; ancestor != null; ancestor = ancestor.ParentElement) {
            if (stopAtContainingBlock && ReferenceEquals(ancestor, request.ContainingBlock)) break;
            if (_semanticNodeIds.TryGetValue(ancestor, out int nodeId) && boundaries.ContainsKey(nodeId)) ancestors.Push(nodeId);
        }
        var envelopes = new List<HtmlRenderVisual>();
        var included = new HashSet<HtmlRenderVisual>();
        while (ancestors.Count > 0) {
            foreach (HtmlRenderVisual envelope in boundaries[ancestors.Pop()]) {
                if (included.Add(envelope)) envelopes.Add(envelope);
            }
        }
        return envelopes;
    }

    private HtmlRenderVisual ApplyFixedLegacyClipEnvelopes(HtmlRenderVisual visual, IReadOnlyList<HtmlRenderVisual> envelopes) {
        if (envelopes.Count == 0) return visual;
        // Paint is already placed in the containing space. Reusing ancestor
        // envelopes clips it without applying their transforms a second time.
        OfficeTransform transform = OfficeTransform.Identity;
        foreach (HtmlRenderEffectGroup effect in envelopes.OfType<HtmlRenderEffectGroup>()) transform = effect.Transform.Then(transform);
        IReadOnlyList<HtmlRenderVisual> content = ReplaceDescendantFormFieldsForPaintEffect(
            new[] { visual }, "ancestor-clip=" + envelopes.OfType<HtmlRenderClipGroup>().First().Source);
        if (!transform.TryInvert(out OfficeTransform inverse)) return visual.CopyStackingContextTo(new HtmlRenderClipGroup(
            visual.X, visual.Y, 0D, 0D, true, true, content, visual.PaintOrder, visual.Source));
        if (inverse != OfficeTransform.Identity) content = new[] { new HtmlRenderEffectGroup(visual.X, visual.Y,
            visual.Width, visual.Height, inverse, 1D, content, 0, visual.Source) };
        for (int index = envelopes.Count - 1; index >= 0; index--) {
            HtmlRenderVisual envelope = envelopes[index];
            content = new[] { envelope switch {
                HtmlRenderClipGroup clip => clip.ProjectPaint(content, 0D, 0D, 0),
                HtmlRenderPathClipGroup path => path.ProjectPaint(content, 0D, 0D, 0),
                HtmlRenderEffectGroup effect => effect.ProjectPaint(content, 0D, 0D, 0),
                _ => throw new InvalidOperationException("Unsupported fixed clipping boundary.")
            } };
        }
        return visual.CopyStackingContextTo(content[0].Translate(0D, 0D, visual.PaintOrder));
    }

    private void CaptureFixedLegacyClipBoundaries(IEnumerable<HtmlRenderVisual> visuals, List<HtmlRenderVisual> path,
        Dictionary<int, IReadOnlyList<HtmlRenderVisual>>? boundaries = null) {
        boundaries ??= _fixedLegacyClipBoundaries;
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
                boundaries[nodeId] = boundary;
            }
            IReadOnlyList<HtmlRenderVisual>? children = GetGroupChildren(visual);
            if (children != null) CaptureFixedLegacyClipBoundaries(children, path, boundaries);
            if (envelope) path.RemoveAt(path.Count - 1);
        }
    }
}
