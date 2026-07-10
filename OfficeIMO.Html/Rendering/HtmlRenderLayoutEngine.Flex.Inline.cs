using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddInlineFlexRun(
        IElement element,
        double availableWidth,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        HtmlRenderBoxStyle inlineStyle,
        string? link,
        double inheritedPaintOffsetX,
        double inheritedPaintOffsetY,
        ICollection<HtmlInlineRun> runs) {
        HtmlRenderBoxStyle flexStyle = BlockifyFlexItemStyle(inlineStyle);
        double outerWidth = ResolveInlineFlexWidth(element, availableWidth, flexStyle, depth + 1);
        if (!flexStyle.ExplicitWidth.HasValue) {
            flexStyle = flexStyle.Clone();
            double targetBoxWidth = Math.Max(0.01D, outerWidth - flexStyle.MarginLeft - flexStyle.MarginRight);
            flexStyle.ExplicitWidth = flexStyle.BorderBox
                ? targetBoxWidth
                : Math.Max(0.01D, targetBoxWidth - flexStyle.HorizontalInsets);
        }

        HtmlRenderFlowBlock atomic = LayoutElement(element, outerWidth, flexStyle, parentStyle, depth + 1);
        runs.Add(new HtmlInlineRun(
            atomic,
            inlineStyle,
            link,
            HtmlRenderStyleResolver.DescribeSource(element),
            inheritedPaintOffsetX,
            inheritedPaintOffsetY));
    }

    private double ResolveInlineFlexWidth(IElement element, double availableWidth, HtmlRenderBoxStyle style, int depth) {
        double availableOuterWidth = Math.Max(1D, availableWidth);
        double availableBoxWidth = Math.Max(1D, availableOuterWidth - style.MarginLeft - style.MarginRight);
        if (style.ExplicitWidth.HasValue) {
            return Math.Min(availableOuterWidth, style.MarginLeft + ResolveBoxWidth(availableBoxWidth, style) + style.MarginRight);
        }

        if (!TryCollectFlexItems(element, availableOuterWidth, style, depth, out List<FlexItem> items)) return availableOuterWidth;
        List<FlexItem> ordered = items.OrderBy(item => item.Style.Order).ThenBy(item => item.SourceIndex).ToList();
        double intrinsicContentWidth;
        if (style.FlexDirection == "column" || style.FlexDirection == "column-reverse") {
            intrinsicContentWidth = ordered.Count == 0 ? 1D : ordered.Max(item => ResolveColumnFlexCrossBasis(item, availableBoxWidth));
        } else {
            foreach (FlexItem item in ordered) item.Basis = ResolveFlexBasis(item, availableBoxWidth);
            intrinsicContentWidth = ordered.Sum(item => item.Basis) + style.ColumnGap * Math.Max(0, ordered.Count - 1);
        }

        double intrinsicBoxWidth = intrinsicContentWidth + style.HorizontalInsets;
        var intrinsicStyle = style.Clone();
        intrinsicStyle.ExplicitWidth = intrinsicStyle.BorderBox ? intrinsicBoxWidth : intrinsicContentWidth;
        double resolvedBoxWidth = ResolveBoxWidth(availableBoxWidth, intrinsicStyle);
        return Math.Max(1D, Math.Min(availableOuterWidth, style.MarginLeft + resolvedBoxWidth + style.MarginRight));
    }
}
