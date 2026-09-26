using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool TryRelayoutBlockForDeferredFloat(
        HtmlRenderFlowBlock block,
        HtmlCssPageGeometry geometry,
        double remainingHeight,
        double pageHeight,
        out HtmlRenderFlowBlock reflowed) {
        reflowed = block;
        if (block.OwnerElement == null || block.Height <= remainingHeight + 0.0001D
            || remainingHeight <= 0.0001D || pageHeight <= 0D || HasInternalForcedBreak(block)) return false;
        IElement root = _document.Body ?? _document.DocumentElement ?? block.OwnerElement;
        bool isRoot = ReferenceEquals(block.OwnerElement, root);
        if (!isRoot && !ReferenceEquals(block.OwnerElement.ParentElement, root)) return false;
        HtmlRenderBoxStyle rootStyle = _styleResolver.Resolve(root, geometry.ContentWidth);
        HtmlRenderBoxStyle style = isRoot
            ? rootStyle
            : _styleResolver.Resolve(block.OwnerElement, geometry.ContentWidth, rootStyle);
        if (!ContainsFloatingDescendant(block.OwnerElement, geometry.ContentWidth, style, isRoot ? 0 : 1)) return false;

        // If the legal break moves a page-sized scroll container as a unit,
        // keep its float at its original local position. Deferring that float
        // while also moving the container would shift it twice on the next page.
        double entryBreak = FindFragmentEnd(block, 0D, remainingHeight, fullPageHeight: pageHeight);
        if (entryBreak > 0.0001D && block.InlineBreakProgress.Any(progress =>
                progress.IsBlockEntry
                && Math.Abs(progress.Offset - entryBreak) <= 0.0001D
                && progress.OwnerElement != null
                && _layoutStyles.TryGetValue(progress.OwnerElement, out HtmlRenderBoxStyle? entryStyle)
                && entryStyle.AvoidBreakInside
                && ContainsFloatingDescendant(progress.OwnerElement, geometry.ContentWidth, entryStyle, 1))) {
            return false;
        }

        try {
            double boundaryHeight = remainingHeight;
            bool deferred = false;
            double leadingAdjustment = isRoot ? 0D : block.LeadingFlowAdjustment;
            for (int pass = 0; pass < 3; pass++) {
                _pagedFloatDeferredInRelayout = false;
                var boundary = new PagedFloatBoundary(boundaryHeight + leadingAdjustment, pageHeight);
                HtmlRenderFlowBlock candidate = isRoot
                    ? LayoutRootElement(root, geometry.ContentWidth, rootStyle, pageBoundary: boundary)
                    : LayoutElement(block.OwnerElement, geometry.ContentWidth, style, rootStyle, 1,
                        pageBoundary: boundary).AdjustLeadingFlowSpace(leadingAdjustment);
                if (!_pagedFloatDeferredInRelayout) break;
                // If widows or orphans forbid a fragment here, pagination moves
                // the entire original block to the next page. Do not retain a
                // float offset computed for the previous page's empty space.
                double fragmentEnd = FindFragmentEnd(candidate, 0D, remainingHeight, fullPageHeight: pageHeight);
                if (fragmentEnd <= 0.0001D || fragmentEnd > boundaryHeight + 0.0001D) break;
                deferred = true;
                reflowed = candidate;
                // A legal line break can precede the physical page edge. Move the
                // deferred float to that actual break so it starts at the next page top.
                if (fragmentEnd >= boundaryHeight - 0.0001D) break;
                boundaryHeight = fragmentEnd;
            }
            return deferred;
        } finally {
            _pagedFloatDeferredInRelayout = false;
        }
    }
}
