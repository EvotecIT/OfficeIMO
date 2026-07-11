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
        HtmlRenderBoxStyle blockStyle = BlockifyFlexItemStyle(inlineStyle);
        double outerWidth = Math.Min(
            Math.Max(1D, availableWidth),
            ResolvePositionedOuterWidth(element, blockStyle, availableWidth, null, null));
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
