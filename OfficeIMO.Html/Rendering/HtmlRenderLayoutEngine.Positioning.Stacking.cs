namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void BuildRootStackingPaintOrders(IEnumerable<HtmlRenderFlowBlock> blocks) {
        var contexts = blocks
            .Where(block => block.StackingZIndex.HasValue)
            .Select(block => new RootStackingContext(block.StackingZIndex!.Value, block.StackingSourceOrder))
            .Concat(_rootPositionedElements.Select(request => new RootStackingContext(request.ZIndex, request.SourceOrder)))
            .Concat(_fixedPositionedElements.Select(request => new RootStackingContext(request.ZIndex, request.SourceOrder)))
            .GroupBy(context => context.SourceOrder)
            .Select(group => group.First())
            .ToList();
        _rootStackingPaintOrders.Clear();
        int negativeOrder = -1000000000;
        foreach (RootStackingContext context in contexts
            .Where(context => context.ZIndex < 0)
            .OrderBy(context => context.ZIndex)
            .ThenBy(context => context.SourceOrder)) {
            _rootStackingPaintOrders[context.SourceOrder] = negativeOrder++;
        }
        int nonNegativeOrder = 1000000000;
        foreach (RootStackingContext context in contexts
            .Where(context => context.ZIndex >= 0)
            .OrderBy(context => context.ZIndex)
            .ThenBy(context => context.SourceOrder)) {
            _rootStackingPaintOrders[context.SourceOrder] = nonNegativeOrder++;
        }
    }

    private int ResolveRootStackingPaintOrder(int sourceOrder, int fallback) =>
        _rootStackingPaintOrders.TryGetValue(sourceOrder, out int paintOrder) ? paintOrder : fallback;

    private static void AppendFlowPaintLayers(ICollection<HtmlRenderVisual> visuals, IEnumerable<FlowPaintLayer> layers) {
        foreach (FlowPaintLayer layer in OrderFlowPaintLayers(layers)) {
            foreach (HtmlRenderVisual visual in layer.Block.Visuals) {
                visuals.Add(visual.Translate(layer.X, layer.Y, visuals.Count));
            }
        }
    }

    private static IEnumerable<FlowPaintLayer> OrderFlowPaintLayers(IEnumerable<FlowPaintLayer> layers) {
        List<FlowPaintLayer> materialized = layers.ToList();
        return materialized
            .Where(layer => layer.Block.StackingZIndex < 0)
            .OrderBy(layer => layer.Block.StackingZIndex)
            .ThenBy(layer => layer.SourceOrder)
            .Concat(materialized
                .Where(layer => !layer.Block.StackingZIndex.HasValue)
                .OrderBy(layer => layer.SourceOrder))
            .Concat(materialized
                .Where(layer => layer.Block.StackingZIndex >= 0)
                .OrderBy(layer => layer.Block.StackingZIndex)
                .ThenBy(layer => layer.SourceOrder));
    }

    private sealed class FlowPaintLayer {
        internal FlowPaintLayer(HtmlRenderFlowBlock block, double x, double y, int sourceOrder) {
            Block = block;
            X = x;
            Y = y;
            SourceOrder = sourceOrder;
        }
        internal HtmlRenderFlowBlock Block { get; }
        internal double X { get; }
        internal double Y { get; }
        internal int SourceOrder { get; }
    }

    private sealed class RootStackingContext {
        internal RootStackingContext(int zIndex, int sourceOrder) {
            ZIndex = zIndex;
            SourceOrder = sourceOrder;
        }
        internal int ZIndex { get; }
        internal int SourceOrder { get; }
    }
}
