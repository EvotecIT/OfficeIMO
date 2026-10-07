using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly Dictionary<IElement, PageFloatEntry> _pageFloatEntries = new();
    private PageFloatPlan _pageFloatPlan = PageFloatPlan.Empty;
    private bool _layingOutPageFloatBody;

    /// <summary>Recognizes the bounded horizontal page-edge float contract; other references retain their reported fallback.</summary>
    private bool TryGetPageFloatSide(IElement element, HtmlRenderBoxStyle style, out string side) {
        side = string.Empty;
        if (_options.Mode != HtmlRenderMode.Paged || _layingOutPageFloatBody || style.WritingMode != "horizontal-tb"
            || !_computedStyles.Elements.TryGetValue(element, out HtmlComputedStyle? computed)
            || !string.Equals(computed.GetValue("float-reference").Trim(), "page", StringComparison.OrdinalIgnoreCase)) return false;
        string value = computed.GetValue("float").Trim().ToLowerInvariant();
        side = value == "block-start" ? "top" : value == "block-end" ? "bottom" : value;
        return side == "top" || side == "bottom" || side == "snap";
    }

    private bool TryAddPageFloatRun(IElement element, HtmlRenderBoxStyle parentStyle, int depth,
        HtmlRenderBoxStyle style, string? link, ICollection<HtmlInlineRun> runs) {
        if (!TryGetPageFloatSide(element, style, out string side)) return false;
        if (!_pageFloatEntries.TryGetValue(element, out PageFloatEntry? entry)) {
            double width = _pageFloatPlan.Placements.TryGetValue(element, out PageFloatPlacement? planned)
                ? planned.Width : _activePageGeometry.ContentWidth;
            HtmlRenderBoxStyle bodyStyle = style.Clone();
            bodyStyle.FloatSide = "none";
            bodyStyle.ClearSide = "none";
            bodyStyle.UnsupportedFloat = string.Empty;
            HtmlRenderFlowBlock body;
            _layingOutPageFloatBody = true;
            try {
                body = LayoutElementWithoutEditableRegionMarker(element, width, bodyStyle, parentStyle, depth + 1).ForPagination();
            } finally {
                _layingOutPageFloatBody = false;
            }
            if (body.Height >= _activePageGeometry.ContentHeight - 0.0001D) {
                ReportUnsupportedFloatValues(element, style);
                return false;
            }
            HtmlRenderBoxStyle regionStyle = style.Clone();
            regionStyle.FloatSide = side;
            body = WrapEditableLayoutRegion(body, element, regionStyle, HtmlRenderLayoutRegionKind.Floating);
            entry = new PageFloatEntry(element, side, body,
                "page-float-anchor:" + GetDocumentOrder(element).ToString(System.Globalization.CultureInfo.InvariantCulture));
            _pageFloatEntries.Add(element, entry);
        }
        // A paint-neutral marker follows the same zero-advance path as bookmark
        // anchors. Its position, rather than the moved figure, selects the page.
        HtmlRenderVisual anchor = new HtmlRenderLogicalTextGroup(string.Empty, 0D, 0D, 0.01D, 0.01D,
            Array.Empty<HtmlRenderVisual>(), _paintOrder++, entry.AnchorSource, isFlowAnchor: true);
        var marker = new HtmlRenderFlowBlock(0D, 0D, new[] { anchor }, HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None, false, entry.AnchorSource);
        HtmlRenderBoxStyle markerStyle = parentStyle.Clone();
        markerStyle.PaintVisible = true;
        runs.Add(new HtmlInlineRun(marker, markerStyle, link, entry.AnchorSource,
            ownerElement: element, isFlowMarker: true));
        return true;
    }

    /// <summary>Reserves whole float boxes using source anchors and monotonic deferral to prevent page-boundary oscillation.</summary>
    private PageFloatPlan ResolvePageFloatPlan(HtmlRenderDocument rendered, HtmlFootnotePagePlan footnotePlan) {
        if (_pageFloatEntries.Count == 0) return PageFloatPlan.Empty;
        var entriesBySource = _pageFloatEntries.Values.ToDictionary(entry => entry.AnchorSource, StringComparer.Ordinal);
        var anchors = new Dictionary<IElement, (int Page, double Y)>();
        foreach (HtmlRenderPage page in rendered.Pages) {
            foreach (HtmlRenderVisual visual in EnumeratePageFloatVisuals(page.Scene)) {
                ChargeLayoutOperation("page-float anchors");
                if (visual.Source != null && entriesBySource.TryGetValue(visual.Source, out PageFloatEntry? entry)
                    && !anchors.ContainsKey(entry.Element)) anchors.Add(entry.Element, (page.PageNumber, visual.Y));
            }
        }
        var placements = new Dictionary<IElement, PageFloatPlacement>();
        var reserved = new Dictionary<int, double>();
        foreach (PageFloatEntry entry in _pageFloatEntries.Values.OrderBy(entry => GetDocumentOrder(entry.Element))) {
            CheckCancellation();
            ChargeLayoutOperation("page-float planning");
            if (!anchors.TryGetValue(entry.Element, out var anchor)) {
                throw new InvalidOperationException("A page float lost its source anchor during pagination.");
            }
            int pageNumber = anchor.Page;
            if (_pageFloatPlan.Placements.TryGetValue(entry.Element, out PageFloatPlacement? previous)) {
                pageNumber = Math.Max(pageNumber, previous.Page);
            }
            while (true) {
                CheckCancellation();
                ChargeLayoutOperation("page-float deferral");
                if (pageNumber > _options.MaxPageCount) throw new InvalidOperationException("Page float deferral exceeded the configured page count.");
                string? pageName = rendered.Pages[Math.Min(pageNumber, rendered.Pages.Count) - 1].PageName;
                HtmlCssPageGeometry geometry = _pageRules.ResolveGeometry(pageNumber, pageName, _options);
                double occupied = reserved.TryGetValue(pageNumber, out double height) ? height : 0D;
                double capacity = geometry.ContentHeight - (footnotePlan.TryGetReservation(pageNumber, out double noteHeight) ? noteHeight : 0D);
                if (entry.Body.Height >= capacity - 0.0001D) {
                    // A deferred left/named page can be smaller than the page
                    // where the float was measured. Try later pages within the
                    // existing cancellation, operation and page-count bounds.
                    pageNumber++;
                    continue;
                }
                if (occupied + entry.Body.Height < capacity - 0.0001D) {
                    double anchorOffset = anchor.Y - geometry.Margins.Top
                        - _pageFloatPlan.ReservedHeight(anchor.Page, "top");
                    string side = entry.Side == "snap"
                        ? anchorOffset < geometry.ContentHeight / 2D ? "top" : "bottom"
                        : entry.Side;
                    placements.Add(entry.Element, new PageFloatPlacement(pageNumber, side, entry.Body.Height, geometry.ContentWidth));
                    reserved[pageNumber] = occupied + entry.Body.Height;
                    break;
                }
                pageNumber++;
            }
        }
        return new PageFloatPlan(placements);
    }

    private double ResolvePageBodyTop(int pageNumber, HtmlCssPageGeometry geometry) =>
        geometry.Margins.Top + _pageFloatPlan.ReservedHeight(pageNumber, "top");

    private void AddPageFloatVisuals(List<HtmlRenderVisual> target, int pageNumber, HtmlCssPageGeometry geometry) {
        if (_pageFloatPlan.Placements.Count == 0) return;
        var floatVisualIndices = new HashSet<int>();
        double top = geometry.Margins.Top;
        double bottom = geometry.Height - geometry.Margins.Bottom - ResolveFootnoteReservation(pageNumber)
            - _pageFloatPlan.ReservedHeight(pageNumber, "bottom");
        foreach (var pair in _pageFloatPlan.Placements.OrderBy(pair => GetDocumentOrder(pair.Key))) {
            CheckCancellation();
            ChargeLayoutOperation("page-float paint");
            PageFloatPlacement placement = pair.Value;
            if (placement.Page != pageNumber || !_pageFloatEntries.TryGetValue(pair.Key, out PageFloatEntry? entry)) continue;
            double y = placement.Side == "top" ? top : bottom;
            int firstVisual = target.Count;
            AddTranslatedVisuals(target, entry.Body.Visuals, geometry.Margins.Left, y, entry.Body);
            for (int index = firstVisual; index < target.Count; index++) floatVisualIndices.Add(index);
            RecordRunningStringAssignments(entry.Body, 0D, entry.Body.Height, y);
            if (placement.Side == "top") top += entry.Body.Height;
            else bottom += entry.Body.Height;
        }
        PreservePageFloatLogicalOwnership(target, floatVisualIndices);
    }

    private static IEnumerable<HtmlRenderVisual> EnumeratePageFloatVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderSemanticGroup semantic ? semantic.Visuals
                : visual is HtmlRenderLogicalTextGroup logical ? logical.Visuals
                : visual is HtmlRenderClipGroup clip ? clip.Visuals
                : visual is HtmlRenderPathClipGroup pathClip ? pathClip.Visuals
                : visual is HtmlRenderEffectGroup effect ? effect.Visuals
                : visual is HtmlRenderLayoutRegion region ? region.Visuals
                : visual is HtmlRenderFormField form ? form.Visuals : null;
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumeratePageFloatVisuals(children)) yield return child;
        }
    }
}
