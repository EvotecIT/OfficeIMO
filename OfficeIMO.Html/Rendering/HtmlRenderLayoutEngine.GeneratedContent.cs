using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddGeneratedInlineRun(
        IElement element,
        HtmlPseudoElementKind kind,
        double width,
        double? containingHeight,
        HtmlRenderBoxStyle parentStyle,
        string? link,
        double inheritedPaintOffsetX,
        double inheritedPaintOffsetY,
        ICollection<HtmlInlineRun> runs) {
        if (!_generatedContent.TryGet(element, kind, out string content)
            || !_styleResolver.TryResolvePseudo(element, kind, width, parentStyle, out HtmlRenderBoxStyle style)
            || style.Display == "none") {
            return;
        }

        string source = DescribePseudoSource(element, kind);
        ReportUnsupportedGeneratedLayout(style, source);
        ResolvePositionPaintOffset(style, width, containingHeight, source, out double offsetX, out double offsetY);
        runs.Add(new HtmlInlineRun(
            ApplyTextTransform(content, style.TextTransform),
            style,
            link,
            source,
            inheritedPaintOffsetX + offsetX,
            inheritedPaintOffsetY + offsetY,
            element));
    }

    private void AddGeneratedContentBlock(
        ICollection<HtmlRenderFlowBlock> blocks,
        IElement element,
        HtmlPseudoElementKind kind,
        double containingWidth,
        HtmlRenderBoxStyle parentStyle) {
        if (!_generatedContent.TryGet(element, kind, out string content)
            || !_styleResolver.TryResolvePseudo(element, kind, containingWidth, parentStyle, out HtmlRenderBoxStyle style)
            || style.Display == "none") {
            return;
        }

        string source = DescribePseudoSource(element, kind);
        ReportUnsupportedGeneratedLayout(style, source);
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        string? link = string.Equals(element.TagName, "a", StringComparison.OrdinalIgnoreCase)
            ? ResolveSafeLink(element.GetAttribute("href"), element)
            : null;
        var run = new HtmlInlineRun(ApplyTextTransform(content, style.TextTransform), style, link, source, ownerElement: element);
        HtmlInlineLayout inline = LayoutInlineRuns(new[] { run }, contentWidth, style);
        double boxHeight = ResolveBoxHeight(inline.Height, style);
        double outerHeight = Math.Max(0.01D, style.MarginTop + boxHeight + style.MarginBottom);
        var visuals = new List<HtmlRenderVisual>();
        bool paintsBlockBox = style.Display == "block" || style.Display == "flow-root" || style.Display == "list-item";
        if (paintsBlockBox) AddGeneratedBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element, source);
        double contentX = style.MarginLeft + style.BorderWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderWidth + style.PaddingTop;
        foreach (HtmlRenderVisual visual in inline.Visuals) {
            visuals.Add(visual.Translate(contentX, contentY, visuals.Count));
        }
        if (paintsBlockBox) AddGeneratedBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element, source);

        IEnumerable<double> breakOffsets = inline.BreakOffsets
            .Select(offset => contentY + offset)
            .Concat(new[] { outerHeight });
        var block = new HtmlRenderFlowBlock(
            containingWidth,
            outerHeight,
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            style.AvoidBreakInside,
            source,
            breakOffsets,
            inline.BreakOffsets.Select(offset => contentY + offset),
            style.Orphans,
            style.Widows,
            pageName: style.PageName ?? parentStyle.PageName);
        blocks.Add(ApplyPositioning(block, style, containingWidth, ResolveContainingBlockHeight(parentStyle), source));
    }

    private void AddGeneratedBoxPaint(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        IElement element,
        string source) {
        double cornerRadius = ResolveBoxCornerRadius(style, width, height, element, source);
        AddBoxShadow(visuals, style, x, y, width, height, cornerRadius, element, source);
        AddBoxBackgroundCore(visuals, style, x, y, width, height, style.BorderWidth, cornerRadius, element, source, source);
        AddBorderPaint(visuals, style, x, y, width, height, cornerRadius, element, source);
    }

    private void AddGeneratedBoxOutlinePaint(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        IElement element,
        string source) {
        double cornerRadius = ResolveBoxCornerRadius(style, width, height, element, source);
        AddOutlinePaint(visuals, style, x, y, width, height, cornerRadius, element, source);
    }

    private static string DescribePseudoSource(IElement element, HtmlPseudoElementKind kind) =>
        HtmlRenderStyleResolver.DescribeSource(element)
        + (kind == HtmlPseudoElementKind.Before ? "::before" : "::after");

    private void ReportUnsupportedGeneratedLayout(HtmlRenderBoxStyle style, string source) {
        string display = style.Display;
        if (display == "flex" || display == "inline-flex") {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.FlexLayoutPending, "Flex layout is not yet active for generated content; text uses normal flow.", HtmlDiagnosticSeverity.Warning, source);
        } else if (display == "grid" || display == "inline-grid") {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GridLayoutPending, "Grid layout is not yet active for generated content; text uses normal flow.", HtmlDiagnosticSeverity.Warning, source);
        }
    }
}
