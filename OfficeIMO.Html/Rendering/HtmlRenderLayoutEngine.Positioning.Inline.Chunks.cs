namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private sealed partial class InlineStackingNode {
        private IReadOnlyList<InlinePaintCapture>? _paintChunks;
        private InlineContainingRect?[]? _chunkBounds;
        private int _firstDecorationChunk = -1;
        private int _lastDecorationChunk = -1;

        // A plain inline descendant participates in each fragment's foreground.
        // A positioned or effected child keeps its existing whole-layer order.
        private bool PaintsWithParentChunk => _paintChunks != null
            && _style.Position == "static" && _style.UsedTransform == "none"
            && _style.UsedIndividualScale == "none" && _style.ClipPath == "none"
            && _style.UnsupportedOpacity.Length == 0 && _style.Opacity >= 1D;

        internal void SetPaintChunks(IReadOnlyList<InlinePaintCapture> chunks) {
            _paintChunks = chunks;
            _chunkBounds = new InlineContainingRect?[chunks.Count];
            if (Bounds == null) return;
            for (int index = 0; index < chunks.Count; index++) {
                if (!chunks[index].Bounds.TryGetValue(Element, out InlineContainingBounds? bounds)) continue;
                InlineContainingRect rect = bounds.ToRect(Bounds.FormattingContainer, isContinuation: false);
                _chunkBounds[index] = rect;
                if (rect.Fragments.Count == 0) continue;
                if (_firstDecorationChunk < 0) _firstDecorationChunk = index;
                _lastDecorationChunk = index;
            }
        }

        /// <summary>
        /// Preserves each ordinary interrupted fragment's background/foreground
        /// sequence inside one owner effect envelope. Later fragment backgrounds
        /// can cover earlier ink without compositing opacity or transforms twice.
        /// </summary>
        private IReadOnlyList<HtmlRenderVisual> ResolveChunkedVisuals() {
            var visuals = new List<HtmlRenderVisual>();
            var outlines = new List<HtmlRenderVisual>();
            List<InlinePaintLayer> wholeLayers = EnumerateWholeLayers().ToList();
            foreach (InlinePaintLayer layer in wholeLayers.Where(layer => layer.ZIndex < 0)
                .OrderBy(layer => layer.ZIndex).ThenBy(layer => layer.SourceOrder)) {
                visuals.AddRange(layer.ResolveVisuals());
            }
            for (int index = 0; index < _paintChunks!.Count; index++) {
                visuals.AddRange(ResolveChunkVisuals(index, outlines));
            }
            foreach (InlinePaintLayer layer in wholeLayers.Where(layer => layer.ZIndex >= 0)
                .OrderBy(layer => layer.ZIndex).ThenBy(layer => layer.SourceOrder)) {
                visuals.AddRange(layer.ResolveVisuals());
            }
            visuals.AddRange(outlines);
            return _owner.ApplyInlineElementPaintEffects(Element, _style, Bounds,
                visuals.Select((visual, index) => visual.Translate(0D, 0D, index)).ToArray(), decorate: false);
        }

        // Plain ancestors paint with each fragment, but their positioned or
        // effected descendants still need one whole layer in the enclosing
        // owner. Resolving that layer here preserves its effect envelope once.
        private IEnumerable<InlinePaintLayer> EnumerateWholeLayers() {
            foreach (InlinePaintLayer layer in Children) {
                if (layer.Node is InlineStackingNode node && node.PaintsWithParentChunk) {
                    foreach (InlinePaintLayer descendant in node.EnumerateWholeLayers()) yield return descendant;
                } else {
                    yield return layer;
                }
            }
        }

        private IReadOnlyList<HtmlRenderVisual> ResolveChunkVisuals(int index,
            ICollection<HtmlRenderVisual> outlines) {
            InlinePaintCapture chunk = _paintChunks![index];
            IReadOnlyList<HtmlRenderVisual> regular = chunk.Owned.TryGetValue(Element, out List<HtmlRenderVisual>? owned)
                ? owned : Array.Empty<HtmlRenderVisual>();
            IEnumerable<InlinePaintLayer> layers = Children.Where(layer => layer.Node?.PaintsWithParentChunk == true)
                .Select(layer => new InlinePaintLayer(layer.ZIndex, layer.SourceOrder,
                    layer.Node!.ResolveChunkVisuals(index, outlines)));
            IReadOnlyList<HtmlRenderVisual> foreground = ComposeInlineLayers(regular, layers);
            InlineContainingRect? bounds = _chunkBounds?[index];
            if (bounds == null || bounds.Fragments.Count == 0) return foreground;
            bool clone = string.Equals(_style.BoxDecorationBreak, "clone", StringComparison.Ordinal);
            HtmlRenderBoxStyle style = clone ? _style : CreateInlineFragmentPaintStyle(_style,
                index == _firstDecorationChunk && Bounds?.IsContinuation != true,
                index == _lastDecorationChunk);
            return _owner.ApplyInlineBoxDecoration(Element, style, bounds, foreground, outlines);
        }
    }
}
