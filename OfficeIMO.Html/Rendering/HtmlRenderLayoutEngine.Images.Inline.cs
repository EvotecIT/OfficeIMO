using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddInlineImageRun(
        IElement element,
        HtmlRenderBoxStyle style,
        string? link,
        double paintOffsetX,
        double paintOffsetY,
        ICollection<HtmlInlineRun> runs) {
        double outerWidth = ResolveFloatingImageOuterWidth(element, style);
        HtmlRenderFlowBlock atomic = LayoutImage(element, outerWidth, style, link);
        runs.Add(new HtmlInlineRun(
            atomic,
            style,
            null,
            HtmlRenderStyleResolver.DescribeSource(element),
            paintOffsetX,
            paintOffsetY,
            element,
            isReplacedImage: true));
    }
}
