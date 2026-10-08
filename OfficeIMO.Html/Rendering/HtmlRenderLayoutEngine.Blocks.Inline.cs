using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddInlineBlockRun(
        IElement element,
        double availableWidth,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        HtmlRenderBoxStyle inlineStyle,
        string? link,
        double inheritedPaintOffsetX,
        double inheritedPaintOffsetY,
        ICollection<HtmlInlineRun> runs) {
        HtmlRenderBoxStyle blockStyle = BlockifyFlexItemStyle(ResolveOrdinaryIntrinsicWidths(element, inlineStyle, availableWidth, depth));
        double outerWidth = ResolvePositionedOuterWidth(element, blockStyle, availableWidth, null, null, depth);
        if (!blockStyle.HasIntrinsicWidths) outerWidth = Math.Min(Math.Max(1D, availableWidth), outerWidth);
        if (!blockStyle.ExplicitWidth.HasValue) {
            blockStyle = blockStyle.Clone();
            SetPositionedExplicitWidth(blockStyle, outerWidth);
        }

        HtmlRenderFlowBlock atomic = LayoutElement(element, outerWidth, blockStyle, parentStyle, depth + 1);
        runs.Add(new HtmlInlineRun(
            atomic,
            inlineStyle,
            link,
            HtmlRenderStyleResolver.DescribeSource(element),
            inheritedPaintOffsetX,
            inheritedPaintOffsetY,
            element));
    }
}
