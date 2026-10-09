using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Delays composition until all fragments of the same inline ancestor are known.
    // In particular opacity must composite overlapping descendants only once.
    private sealed class InlinePaintCapture {
        private readonly HtmlRenderLayoutEngine _owner;
        internal InlinePaintCapture(HtmlRenderLayoutEngine owner) => _owner = owner;
        internal List<HtmlRenderVisual> Visuals { get; } = new();
        internal Dictionary<IElement, List<HtmlRenderVisual>> Owned { get; } = new();
        internal Dictionary<IElement, InlineContainingBounds> Bounds { get; } = new();

        internal void Capture(IEnumerable<HtmlRenderVisual> visuals,
            IReadOnlyDictionary<IElement, List<HtmlRenderVisual>> owned,
            IReadOnlyDictionary<IElement, InlineContainingBounds> bounds) {
            Visuals.AddRange(visuals);
            foreach (var item in owned) Owned.Add(item.Key, item.Value);
            foreach (var item in bounds) Bounds.Add(item.Key, item.Value);
        }

        internal void Merge(InlinePaintCapture other, double offsetX, double offsetY) {
            foreach (HtmlRenderVisual visual in other.Visuals) Visuals.Add(visual.Translate(offsetX, offsetY, Visuals.Count));
            foreach (var item in other.Owned) {
                if (!Owned.TryGetValue(item.Key, out List<HtmlRenderVisual>? destination)) {
                    destination = new List<HtmlRenderVisual>();
                    Owned.Add(item.Key, destination);
                }
                foreach (HtmlRenderVisual visual in item.Value) destination.Add(visual.Translate(offsetX, offsetY, destination.Count));
            }
            foreach (var item in other.Bounds) {
                if (!Bounds.TryGetValue(item.Key, out InlineContainingBounds? destination)) {
                    destination = new InlineContainingBounds(_owner);
                    Bounds.Add(item.Key, destination);
                }
                destination.Merge(item.Value, offsetX, offsetY);
            }
        }
    }
}
