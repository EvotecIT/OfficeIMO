using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private sealed class PageFloatEntry {
        internal PageFloatEntry(IElement element, string side, HtmlRenderFlowBlock body, string anchorSource) {
            Element = element;
            Side = side;
            Body = body;
            AnchorSource = anchorSource;
        }
        internal IElement Element { get; }
        internal string Side { get; }
        internal HtmlRenderFlowBlock Body { get; }
        internal string AnchorSource { get; }
    }

    private sealed class PageFloatPlacement {
        internal PageFloatPlacement(int page, string side, double height, double width) {
            Page = page;
            Side = side;
            Height = height;
            Width = width;
        }
        internal int Page { get; }
        internal string Side { get; }
        internal double Height { get; }
        internal double Width { get; }
    }

    private sealed class PageFloatPlan {
        internal static readonly PageFloatPlan Empty = new(new Dictionary<IElement, PageFloatPlacement>());
        private readonly Dictionary<(int Page, string Side), double> _reservations = new();
        internal PageFloatPlan(IReadOnlyDictionary<IElement, PageFloatPlacement> placements) {
            Placements = placements;
            foreach (PageFloatPlacement placement in placements.Values) {
                var key = (placement.Page, placement.Side);
                _reservations[key] = ReservedHeight(key.Page, key.Side) + placement.Height;
                MaximumPageNumber = Math.Max(MaximumPageNumber, placement.Page);
            }
        }
        internal IReadOnlyDictionary<IElement, PageFloatPlacement> Placements { get; }
        internal int MaximumPageNumber { get; }
        internal double ReservedHeight(int page, string side) =>
            _reservations.TryGetValue((page, side), out double height) ? height : 0D;
        internal bool EquivalentTo(PageFloatPlan other) => Placements.Count == other.Placements.Count
            && Placements.All(pair => other.Placements.TryGetValue(pair.Key, out PageFloatPlacement? placement)
                && placement.Page == pair.Value.Page && placement.Side == pair.Value.Side
                && Math.Abs(placement.Height - pair.Value.Height) <= 0.0001D
                && Math.Abs(placement.Width - pair.Value.Width) <= 0.0001D);
    }
}
