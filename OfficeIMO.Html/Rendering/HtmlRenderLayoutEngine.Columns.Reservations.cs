using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IElement? _activeColumnOwner;

    private IReadOnlyList<HtmlRenderFlowBlock> BuildColumnChildren(
        IElement element, double width, HtmlRenderBoxStyle style, int depth) {
        IElement? previous = _activeColumnOwner;
        double previousWidth = _activeColumnWidth;
        double previousHeight = _activeColumnHeight;
        _activeColumnOwner = _options.Mode == HtmlRenderMode.Paged && style.WritingMode == "horizontal-tb" ? element : null;
        _activeColumnWidth = width;
        _activeColumnHeight = Math.Min(ResolveDeclaredColumnContentHeight(style) ?? _activePageGeometry.ContentHeight, _activePageGeometry.ContentHeight);
        try {
            return BuildMultiColumnChildBlocks(element, element.ChildNodes, width, style, depth);
        } finally {
            _activeColumnOwner = previous;
            _activeColumnWidth = previousWidth;
            _activeColumnHeight = previousHeight;
        }
    }

    private IElement? ResolveColumnReferenceOwner(IElement element) {
        if (_activeColumnOwner == null || _layingOutColumnEdgeFloatBody || _layingOutPageFloatBody || !_computedStyles.Elements.TryGetValue(element, out HtmlComputedStyle? computed)
            || !string.Equals(computed.GetValue("float-reference").Trim(), "column", StringComparison.OrdinalIgnoreCase)) return null;
        for (IElement? parent = element.ParentElement; parent != null; parent = parent.ParentElement) {
            if (ReferenceEquals(parent, _activeColumnOwner)) return parent;
            if (_computedStyles.Elements.TryGetValue(parent, out HtmlComputedStyle? parentStyle)) {
                // Extracted note bodies are outside the ancestor column flow;
                // their descendant figures cannot leave anchors in that owner.
                if (string.Equals(parentStyle.GetValue("float").Trim(), "footnote", StringComparison.OrdinalIgnoreCase)) return null;
                string count = parentStyle.GetValue("column-count").Trim();
                string width = parentStyle.GetValue("column-width").Trim();
                if ((count.Length > 0 && count != "auto") || (width.Length > 0 && width != "auto")) return null;
            }
        }
        return null;
    }

    /// <summary>
    /// Reflows column bodies against shared edge-figure and note reservations.
    /// Monotonic placement and a bounded pass count prevent boundary oscillation.
    /// </summary>
    private MultiColumnPlan ResolveColumnReservedLayout(
        IElement owner, IReadOnlyList<HtmlRenderFlowBlock> children, MultiColumnPlan body,
        double width, double height) {
        HtmlFootnoteEntry[] entries = _footnoteEntries.Values.Where(e => ReferenceEquals(e.ColumnOwner, owner))
            .OrderBy(e => e.Number).ToArray();
        ColumnEdgeFloatEntry[] floatEntries = _columnEdgeFloatEntries.Values.Where(e => ReferenceEquals(e.Owner, owner))
            .OrderBy(e => GetDocumentOrder(e.Element)).ToArray();
        if (entries.Length == 0 && floatEntries.Length == 0) return body;
        ColumnEdgeFloatPlan floats = PlanColumnEdgeFloats(body, floatEntries, height, null, null);
        ColumnNotePlan notes = PlanColumnNotes(body, entries, width, height, null, floats);
        int maximumPasses = Math.Min(64, Math.Max(8, entries.Length + floatEntries.Length + 4));
        for (int pass = 0; pass < maximumPasses; pass++) {
            CheckCancellation();
            body = BuildMultiColumnPlan(children, height, _options.MaxColumnCount, throwOnLimit: true, notes, floats);
            ColumnEdgeFloatPlan nextFloats = PlanColumnEdgeFloats(body, floatEntries, height, floats, notes);
            ColumnNotePlan nextNotes = PlanColumnNotes(body, entries, width, height, notes, nextFloats);
            if (nextNotes.EquivalentTo(notes) && nextFloats.EquivalentTo(floats)) {
                body = AppendColumnNoteFragments(body, nextNotes, width, height);
                return AppendColumnEdgeFloatFragments(body, nextFloats, nextNotes, height);
            }
            notes = nextNotes;
            floats = nextFloats;
        }
        throw new HtmlDomLimitException(HtmlRenderDiagnosticCodes.PaginationConvergenceLimitExceeded,
            "Reserved column content exceeded its bounded convergence limit.", "ColumnReservationReflowPasses", maximumPasses + 1, maximumPasses);
    }

}
