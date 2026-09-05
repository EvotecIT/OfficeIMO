using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void RecordInlineStaticMarker(
        HtmlInlineRun run,
        IElement? formattingContainer,
        double x,
        double y,
        double lineHeight,
        IDictionary<IElement, InlineContainingBounds> bounds) {
        if (run.PositionedMarkerElement == null) return;
        IElement? anchorParent = formattingContainer ?? run.OwnerElement;
        if (anchorParent != null) {
            _inlineStaticPositions[run.PositionedMarkerElement] = new InlineStaticPosition(
                anchorParent,
                x + run.PaintOffsetX,
                y + run.PaintOffsetY);
        }
        RecordInlineOwnerGeometry(run, formattingContainer, x, y, 0.01D, lineHeight, bounds);
    }

    private void RecordInlineOwnerGeometry(
        HtmlInlineRun run,
        IElement? formattingContainer,
        double x,
        double y,
        double width,
        double height,
        IDictionary<IElement, InlineContainingBounds> bounds) {
        for (IElement? current = run.OwnerElement; current != null; current = current.ParentElement) {
            if ((_localPositionedElements.ContainsKey(current) || _inlineStackingElements.Contains(current))
                && _layoutStyles.TryGetValue(current, out HtmlRenderBoxStyle? style)
                && style.Display == "inline") {
                if (!bounds.TryGetValue(current, out InlineContainingBounds? currentBounds)) {
                    currentBounds = new InlineContainingBounds();
                    bounds[current] = currentBounds;
                }
                currentBounds.Include(
                    x + run.PaintOffsetX,
                    y + run.PaintOffsetY,
                    Math.Max(0.01D, width),
                    Math.Max(0.01D, height));
            }
            if (ReferenceEquals(current, formattingContainer)) break;
        }
    }

    private void EnsureInlineStackingOwner(
        IElement? ownerElement,
        IElement? formattingContainer,
        IDictionary<IElement, List<HtmlRenderVisual>> ownedVisuals) {
        IElement? stackingElement = FindNearestInlineStackingElement(ownerElement, formattingContainer);
        if (stackingElement != null && !ownedVisuals.ContainsKey(stackingElement)) {
            ownedVisuals[stackingElement] = new List<HtmlRenderVisual>();
        }
    }

    private void AddInlineOwnedVisual(
        ICollection<HtmlRenderVisual> rootVisuals,
        IDictionary<IElement, List<HtmlRenderVisual>> ownedVisuals,
        HtmlRenderVisual visual,
        IElement? ownerElement,
        IElement? formattingContainer) {
        IElement? stackingElement = FindNearestInlineStackingElement(ownerElement, formattingContainer);
        if (stackingElement == null) {
            rootVisuals.Add(visual);
            return;
        }
        if (!ownedVisuals.TryGetValue(stackingElement, out List<HtmlRenderVisual>? visuals)) {
            visuals = new List<HtmlRenderVisual>();
            ownedVisuals[stackingElement] = visuals;
        }
        visuals.Add(visual);
    }

    private IElement? FindNearestInlineStackingElement(IElement? element, IElement? formattingContainer) {
        for (IElement? current = element; current != null; current = current.ParentElement) {
            if (_inlineStackingElements.Contains(current)) return current;
            if (ReferenceEquals(current, formattingContainer)) break;
        }
        return null;
    }

    private IReadOnlyList<HtmlRenderVisual> ComposeInlinePositionedVisuals(
        IReadOnlyList<HtmlRenderVisual> content,
        IReadOnlyDictionary<IElement, List<HtmlRenderVisual>> ownedVisuals,
        IReadOnlyDictionary<IElement, InlineContainingBounds> bounds,
        IElement? formattingContainer,
        bool isInlineContinuation) {
        if (formattingContainer == null) return content;
        var nodes = new Dictionary<IElement, InlineStackingNode>();
        foreach (IElement element in ownedVisuals.Keys) CreateInlineStackingNode(element, formattingContainer, ownedVisuals, nodes);

        var positionedPlacements = new List<InlinePositionedPlacement>();
        foreach (KeyValuePair<IElement, InlineContainingBounds> entry in bounds) {
            InlineContainingRect rect = entry.Value.ToRect(formattingContainer, isInlineContinuation);
            _inlineContainingRects[entry.Key] = rect;
            if (nodes.TryGetValue(entry.Key, out InlineStackingNode? node)) node.Bounds = rect;
            if (!_localPositionedElements.TryGetValue(entry.Key, out List<PositionedElementRequest>? requests)) continue;
            foreach (PositionedElementRequest request in requests.Where(item => ReferenceEquals(item.ContainingBlock, entry.Key))) {
                positionedPlacements.Add(new InlinePositionedPlacement(request, rect));
            }
        }

        var rootLayers = new List<InlinePaintLayer>();
        foreach (InlineStackingNode node in nodes.Values) {
            var layer = new InlinePaintLayer(node.ZIndex, node.SourceOrder, node);
            if (node.Parent != null) node.Parent.Children.Add(layer);
            else rootLayers.Add(layer);
        }
        foreach (InlinePositionedPlacement placement in positionedPlacements) {
            PositionedLayer positioned = placement.Request.Resolve(this, placement.Rect.Width, placement.Rect.Height);
            var positionedVisuals = positioned.Block.Visuals
                .Select((visual, index) => visual.Translate(placement.Rect.X + positioned.X, placement.Rect.Y + positioned.Y, index))
                .ToList();
            var layer = new InlinePaintLayer(placement.Request.ZIndex, placement.Request.SourceOrder, positionedVisuals);
            IElement? ownerElement = FindNearestInlineStackingElement(placement.Request.ContainingBlock, formattingContainer);
            if (ownerElement != null) {
                InlineStackingNode owner = CreateInlineStackingNode(ownerElement, formattingContainer, ownedVisuals, nodes);
                owner.Children.Add(layer);
            } else {
                rootLayers.Add(layer);
            }
        }
        if (rootLayers.Count == 0) return content;
        return ComposeInlineLayers(content, rootLayers);
    }

    private InlineStackingNode CreateInlineStackingNode(
        IElement element,
        IElement formattingContainer,
        IReadOnlyDictionary<IElement, List<HtmlRenderVisual>> ownedVisuals,
        IDictionary<IElement, InlineStackingNode> nodes) {
        if (nodes.TryGetValue(element, out InlineStackingNode? existing)) return existing;
        HtmlRenderBoxStyle style = _layoutStyles[element];
        var node = new InlineStackingNode(
            this,
            element,
            style,
            ResolvePositionedZIndex(element, style),
            GetPositionedSourceOrder(element),
            ownedVisuals.TryGetValue(element, out List<HtmlRenderVisual>? regular) ? regular : new List<HtmlRenderVisual>());
        nodes[element] = node;
        IElement? parentElement = FindNearestInlineStackingElement(element.ParentElement, formattingContainer);
        if (parentElement != null) node.Parent = CreateInlineStackingNode(parentElement, formattingContainer, ownedVisuals, nodes);
        return node;
    }

    private static IReadOnlyList<HtmlRenderVisual> ComposeInlineLayers(
        IReadOnlyList<HtmlRenderVisual> content,
        IEnumerable<InlinePaintLayer> layers) {
        List<InlinePaintLayer> materialized = layers.ToList();
        var combined = new List<HtmlRenderVisual>();
        foreach (InlinePaintLayer layer in materialized.Where(item => item.ZIndex < 0).OrderBy(item => item.ZIndex).ThenBy(item => item.SourceOrder)) {
            combined.AddRange(layer.ResolveVisuals());
        }
        combined.AddRange(content);
        foreach (InlinePaintLayer layer in materialized.Where(item => item.ZIndex >= 0).OrderBy(item => item.ZIndex).ThenBy(item => item.SourceOrder)) {
            combined.AddRange(layer.ResolveVisuals());
        }
        return combined;
    }

    private sealed class InlineContainingBounds {
        private double _left = double.PositiveInfinity;
        private double _top = double.PositiveInfinity;
        private double _right = double.NegativeInfinity;
        private double _bottom = double.NegativeInfinity;
        private readonly List<InlineFragmentRect> _fragments = new List<InlineFragmentRect>();

        internal void Include(double x, double y, double width, double height) {
            _left = Math.Min(_left, x);
            _top = Math.Min(_top, y);
            _right = Math.Max(_right, x + width);
            _bottom = Math.Max(_bottom, y + height);
            IncludeFragment(x, y, width, height);
        }

        private void IncludeFragment(double x, double y, double width, double height) {
            const double tolerance = 0.01D;
            for (int index = 0; index < _fragments.Count; index++) {
                InlineFragmentRect fragment = _fragments[index];
                if (Math.Abs(fragment.Y - y) > tolerance || Math.Abs(fragment.Height - height) > tolerance) continue;
                double right = x + width;
                if (right < fragment.X - tolerance || x > fragment.Right + tolerance) continue;
                _fragments[index] = new InlineFragmentRect(
                    Math.Min(fragment.X, x),
                    Math.Min(fragment.Y, y),
                    Math.Max(fragment.Right, right) - Math.Min(fragment.X, x),
                    Math.Max(fragment.Bottom, y + height) - Math.Min(fragment.Y, y));
                return;
            }
            _fragments.Add(new InlineFragmentRect(x, y, width, height));
        }

        internal InlineContainingRect ToRect(IElement formattingContainer, bool isContinuation) =>
            new InlineContainingRect(
                formattingContainer,
                double.IsPositiveInfinity(_left) ? 0D : _left,
                double.IsPositiveInfinity(_top) ? 0D : _top,
                double.IsNegativeInfinity(_right) ? 0.01D : _right - _left,
                double.IsNegativeInfinity(_bottom) ? 0.01D : _bottom - _top,
                _fragments.OrderBy(item => item.Y).ThenBy(item => item.X).ToArray(),
                isContinuation);
    }

    private readonly struct InlineFragmentRect {
        internal InlineFragmentRect(double x, double y, double width, double height) {
            X = x;
            Y = y;
            Width = Math.Max(0.01D, width);
            Height = Math.Max(0.01D, height);
        }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal double Right => X + Width;
        internal double Bottom => Y + Height;
    }

    private sealed class InlineContainingRect {
        internal InlineContainingRect(
            IElement formattingContainer,
            double x,
            double y,
            double width,
            double height,
            IReadOnlyList<InlineFragmentRect> fragments,
            bool isContinuation) {
            FormattingContainer = formattingContainer;
            X = x;
            Y = y;
            Width = Math.Max(0.01D, width);
            Height = Math.Max(0.01D, height);
            Fragments = fragments;
            IsContinuation = isContinuation;
        }
        internal IElement FormattingContainer { get; }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal IReadOnlyList<InlineFragmentRect> Fragments { get; }
        internal bool IsContinuation { get; }
    }

    private sealed class InlineStaticPosition {
        internal InlineStaticPosition(IElement formattingContainer, double x, double y) {
            FormattingContainer = formattingContainer;
            X = x;
            Y = y;
        }
        internal IElement FormattingContainer { get; }
        internal double X { get; }
        internal double Y { get; }
    }

    private sealed class InlinePositionedPlacement {
        internal InlinePositionedPlacement(PositionedElementRequest request, InlineContainingRect rect) {
            Request = request;
            Rect = rect;
        }
        internal PositionedElementRequest Request { get; }
        internal InlineContainingRect Rect { get; }
    }

    private sealed class InlineStackingNode {
        private readonly HtmlRenderLayoutEngine _owner;
        private readonly HtmlRenderBoxStyle _style;

        internal InlineStackingNode(HtmlRenderLayoutEngine owner, IElement element, HtmlRenderBoxStyle style, int zIndex, int sourceOrder, IReadOnlyList<HtmlRenderVisual> regularVisuals) {
            _owner = owner;
            Element = element;
            _style = style;
            ZIndex = zIndex;
            SourceOrder = sourceOrder;
            RegularVisuals = regularVisuals;
        }
        internal IElement Element { get; }
        internal int ZIndex { get; }
        internal int SourceOrder { get; }
        internal IReadOnlyList<HtmlRenderVisual> RegularVisuals { get; }
        internal InlineContainingRect? Bounds { get; set; }
        internal InlineStackingNode? Parent { get; set; }
        internal List<InlinePaintLayer> Children { get; } = new List<InlinePaintLayer>();
        internal IReadOnlyList<HtmlRenderVisual> ResolveVisuals() =>
            _owner.ApplyInlineElementPaintEffects(Element, _style, Bounds, ComposeInlineLayers(RegularVisuals, Children));
    }

    private sealed class InlinePaintLayer {
        private readonly InlineStackingNode? _node;
        private readonly IReadOnlyList<HtmlRenderVisual>? _visuals;

        internal InlinePaintLayer(int zIndex, int sourceOrder, InlineStackingNode node) {
            ZIndex = zIndex;
            SourceOrder = sourceOrder;
            _node = node;
        }

        internal InlinePaintLayer(int zIndex, int sourceOrder, IReadOnlyList<HtmlRenderVisual> visuals) {
            ZIndex = zIndex;
            SourceOrder = sourceOrder;
            _visuals = visuals;
        }

        internal int ZIndex { get; }
        internal int SourceOrder { get; }
        internal IReadOnlyList<HtmlRenderVisual> ResolveVisuals() => _node?.ResolveVisuals() ?? _visuals ?? Array.Empty<HtmlRenderVisual>();
    }
}
