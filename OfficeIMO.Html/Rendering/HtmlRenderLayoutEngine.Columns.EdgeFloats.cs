using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly Dictionary<IElement, ColumnEdgeFloatEntry> _columnEdgeFloatEntries = new();
    private double _activeColumnWidth;
    private double _activeColumnHeight;
    private bool _layingOutColumnEdgeFloatBody;

    private bool TryGetColumnEdgeFloatSide(IElement element, HtmlRenderBoxStyle style,
        out string side, out IElement? owner) {
        side = string.Empty;
        owner = null;
        if (_layingOutColumnEdgeFloatBody || style.WritingMode != "horizontal-tb") return false;
        owner = ResolveColumnReferenceOwner(element);
        if (owner == null || !_computedStyles.Elements.TryGetValue(element, out HtmlComputedStyle? computed)) return false;
        string value = computed.GetValue("float").Trim().ToLowerInvariant();
        side = value == "block-start" ? "top" : value == "block-end" ? "bottom" : value;
        return side == "top" || side == "bottom";
    }

    /// <summary>Extracts a whole figure and leaves a zero-advance anchor to identify its originating column.</summary>
    private bool TryAddColumnEdgeFloatRun(IElement element, HtmlRenderBoxStyle parentStyle, int depth,
        HtmlRenderBoxStyle style, string? link, ICollection<HtmlInlineRun> runs) {
        if (!TryGetColumnEdgeFloatSide(element, style, out string side, out IElement? owner)) return false;
        if (!_columnEdgeFloatEntries.TryGetValue(element, out ColumnEdgeFloatEntry? entry)) {
            HtmlRenderBoxStyle bodyStyle = style.Clone();
            bodyStyle.FloatSide = "none";
            bodyStyle.ClearSide = "none";
            bodyStyle.UnsupportedFloat = string.Empty;
            HtmlRenderFlowBlock body;
            _layingOutColumnEdgeFloatBody = true;
            try {
                body = LayoutElementWithoutEditableRegionMarker(element, _activeColumnWidth, bodyStyle,
                    parentStyle, depth + 1).ForPagination();
            } finally {
                _layingOutColumnEdgeFloatBody = false;
            }
            if (body.Height >= _activeColumnHeight - 0.0001D) {
                ReportUnsupportedFloatValues(element, style);
                return false;
            }
            HtmlRenderBoxStyle regionStyle = style.Clone();
            regionStyle.FloatSide = side;
            body = WrapEditableLayoutRegion(body, element, regionStyle, HtmlRenderLayoutRegionKind.Floating);
            // The figure is one atomic reservation even when its caption has
            // internal line breaks. The outer page must not slice that box.
            body = new HtmlRenderFlowBlock(body.Width, body.Height, body.Visuals, body.BreakBefore,
                body.BreakAfter, true, body.Source, new[] { body.Height },
                runningStringAssignments: body.RunningStringAssignments);
            entry = new ColumnEdgeFloatEntry(element, owner!, side, body,
                "column-float-anchor:" + GetDocumentOrder(element).ToString(System.Globalization.CultureInfo.InvariantCulture));
            _columnEdgeFloatEntries.Add(element, entry);
        }
        HtmlRenderVisual anchor = new HtmlRenderLogicalTextGroup(string.Empty, 0D, 0D, 0.01D, 0.01D,
            Array.Empty<HtmlRenderVisual>(), _paintOrder++, entry.Anchor, isFlowAnchor: true);
        var marker = new HtmlRenderFlowBlock(0D, 0D, new[] { anchor }, HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None, false, entry.Anchor);
        HtmlRenderBoxStyle markerStyle = parentStyle.Clone();
        markerStyle.PaintVisible = true;
        runs.Add(new HtmlInlineRun(marker, markerStyle, link, entry.Anchor, ownerElement: element, isFlowMarker: true));
        return true;
    }

    /// <summary>Reserves only the anchor's column; monotonically defers figures when that column cannot contain them.</summary>
    private ColumnEdgeFloatPlan PlanColumnEdgeFloats(MultiColumnPlan body, IReadOnlyList<ColumnEdgeFloatEntry> entries,
        double height, ColumnEdgeFloatPlan? previous, ColumnNotePlan? notes) {
        var placements = new Dictionary<string, ColumnEdgeFloatPlacement>(StringComparer.Ordinal);
        var reservations = new Dictionary<int, double>();
        foreach (ColumnEdgeFloatEntry entry in entries) {
            int column = FindColumnFloatAnchor(body, entry.Anchor, height, out double anchorHeight);
            int anchorColumn = column;
            if (column < 0) throw new InvalidOperationException("A column float lost its source anchor during fragmentation.");
            if (previous != null && previous.Placements.TryGetValue(entry.Anchor, out ColumnEdgeFloatPlacement? old))
                column = Math.Max(column, old.Column);
            while (true) {
                CheckCancellation();
                ChargeLayoutOperation("column float reservation");
                EnsureMultiColumnLimit(column + 1);
                reservations.TryGetValue(column, out double occupied);
                double capacity = height - (notes?.Reserved(column) ?? 0D)
                    - (column == anchorColumn ? anchorHeight : 0D);
                if (occupied + entry.Body.Height <= capacity + 0.0001D) {
                    reservations[column] = occupied + entry.Body.Height;
                    break;
                }
                column++;
            }
            placements.Add(entry.Anchor, new ColumnEdgeFloatPlacement(entry, column));
        }
        return new ColumnEdgeFloatPlan(placements);
    }

    private int FindColumnFloatAnchor(MultiColumnPlan body, string source, double height, out double anchorHeight) {
        anchorHeight = 0D;
        foreach (MultiColumnFragment fragment in body.Fragments) {
            CheckCancellation();
            ChargeLayoutOperation("column float anchor lookup");
            HtmlRenderVisual? anchor = EnumeratePageFloatVisuals(SliceBlockVisuals(fragment.Block, fragment.Start, fragment.End))
                .FirstOrDefault(v => v.Source == source && v is HtmlRenderLogicalTextGroup { IsFlowAnchor: true });
            if (anchor == null) continue;
            if (fragment.Block.Height > 0.0001D) {
                double offset = fragment.Start + anchor.LayoutY;
                double before = FindFragmentEnd(fragment.Block, 0D, Math.Max(0D, offset), offset, fullPageHeight: height);
                anchorHeight = Math.Max(0D, FindNextColumnBreak(fragment.Block, offset) - before);
                if (fragment.Block.AvoidBreakInside) anchorHeight = Math.Max(anchorHeight, fragment.Block.Height);
            }
            return fragment.Column;
        }
        return -1;
    }

    private MultiColumnPlan AppendColumnEdgeFloatFragments(MultiColumnPlan body, ColumnEdgeFloatPlan floats,
        ColumnNotePlan notes, double height) {
        if (floats.Placements.Count == 0) return body;
        var fragments = new List<MultiColumnFragment>(body.Fragments);
        var top = new Dictionary<int, double>();
        var bottom = new Dictionary<int, double>();
        foreach (ColumnEdgeFloatPlacement placement in floats.Placements.Values) {
            CheckCancellation();
            ChargeLayoutOperation("column float painting");
            int column = placement.Column;
            bool atTop = placement.Entry.Side == "top";
            Dictionary<int, double> cursors = atTop ? top : bottom;
            if (!cursors.TryGetValue(column, out double y))
                y = atTop ? 0D : height - notes.Reserved(column) - floats.Bottom(column);
            HtmlRenderFlowBlock block = placement.Entry.Body;
            fragments.Add(new MultiColumnFragment(block, 0D, block.Height, column, y));
            cursors[column] = y + block.Height;
        }
        int count = Math.Max(body.ColumnCount, floats.Placements.Values.Max(p => p.Column) + 1);
        return new MultiColumnPlan(fragments, count, Math.Max(body.UsedHeight, height));
    }

    private sealed record ColumnEdgeFloatEntry(IElement Element, IElement Owner, string Side, HtmlRenderFlowBlock Body, string Anchor);
    private sealed record ColumnEdgeFloatPlacement(ColumnEdgeFloatEntry Entry, int Column);

    private sealed class ColumnEdgeFloatPlan {
        private readonly Dictionary<int, double> _top = new();
        private readonly Dictionary<int, double> _bottom = new();
        internal ColumnEdgeFloatPlan(IReadOnlyDictionary<string, ColumnEdgeFloatPlacement> placements) {
            Placements = placements;
            foreach (ColumnEdgeFloatPlacement placement in placements.Values) {
                Dictionary<int, double> heights = placement.Entry.Side == "top" ? _top : _bottom;
                heights.TryGetValue(placement.Column, out double height);
                heights[placement.Column] = height + placement.Entry.Body.Height;
            }
        }
        internal IReadOnlyDictionary<string, ColumnEdgeFloatPlacement> Placements { get; }
        internal double Top(int column) => _top.TryGetValue(column, out double height) ? height : 0D;
        internal double Bottom(int column) => _bottom.TryGetValue(column, out double height) ? height : 0D;
        internal double Reserved(int column) => Top(column) + Bottom(column);
        internal bool EquivalentTo(ColumnEdgeFloatPlan other) => Placements.Count == other.Placements.Count
            && Placements.All(pair => other.Placements.TryGetValue(pair.Key, out ColumnEdgeFloatPlacement? placement) && placement == pair.Value);
    }
}
