using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Delays composition until all fragments of the same inline ancestor are known.
    // In particular opacity must composite overlapping descendants only once.
    private sealed class InlinePaintCapture {
        private readonly HtmlRenderLayoutEngine _owner;
        internal InlinePaintCapture(HtmlRenderLayoutEngine owner) => _owner = owner;
        internal bool IsOrdinaryText { get; private set; }
        internal List<HtmlRenderVisual> Visuals { get; } = new();
        internal Dictionary<IElement, List<HtmlRenderVisual>> Owned { get; } = new();
        internal Dictionary<IElement, InlineContainingBounds> Bounds { get; } = new();

        internal void Capture(IEnumerable<HtmlRenderVisual> visuals,
            IReadOnlyDictionary<IElement, List<HtmlRenderVisual>> owned,
            IReadOnlyDictionary<IElement, InlineContainingBounds> bounds, bool ordinaryText) {
            _owner.CheckCancellation();
            _owner.ChargeLayoutOperations((long)owned.Count + bounds.Count, "inline paint capture ownership");
            IsOrdinaryText = ordinaryText;
            foreach (HtmlRenderVisual visual in visuals) {
                _owner.CheckCancellation();
                _owner.ChargeLayoutOperation("inline paint capture visuals");
                Visuals.Add(visual);
            }
            foreach (var item in owned) {
                _owner.CheckCancellation();
                Owned.Add(item.Key, item.Value);
            }
            foreach (var item in bounds) {
                _owner.CheckCancellation();
                Bounds.Add(item.Key, item.Value);
            }
        }

        internal void Merge(InlinePaintCapture other, double offsetX, double offsetY) {
            _owner.CheckCancellation();
            _owner.ChargeLayoutOperations((long)other.Visuals.Count + other.Owned.Count + other.Bounds.Count,
                "inline paint capture merging");
            foreach (HtmlRenderVisual visual in other.Visuals) {
                _owner.CheckCancellation();
                Visuals.Add(visual.Translate(offsetX, offsetY, Visuals.Count));
            }
            foreach (var item in other.Owned) {
                _owner.CheckCancellation();
                _owner.ChargeLayoutOperations(item.Value.Count, "inline paint capture owned visuals");
                if (!Owned.TryGetValue(item.Key, out List<HtmlRenderVisual>? destination)) {
                    destination = new List<HtmlRenderVisual>();
                    Owned.Add(item.Key, destination);
                }
                foreach (HtmlRenderVisual visual in item.Value) {
                    _owner.CheckCancellation();
                    destination.Add(visual.Translate(offsetX, offsetY, destination.Count));
                }
            }
            foreach (var item in other.Bounds) {
                _owner.CheckCancellation();
                if (!Bounds.TryGetValue(item.Key, out InlineContainingBounds? destination)) {
                    destination = new InlineContainingBounds(_owner);
                    Bounds.Add(item.Key, destination);
                }
                destination.Merge(item.Value, offsetX, offsetY);
            }
        }
    }
}
