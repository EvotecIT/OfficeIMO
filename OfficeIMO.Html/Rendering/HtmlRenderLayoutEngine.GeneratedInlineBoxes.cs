using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock CreateGeneratedInlineBox(
        IElement element,
        HtmlRenderBoxStyle style,
        HtmlGeneratedContent content,
        string? link,
        string source,
        double containingWidth) {
        var runs = new List<HtmlInlineRun>();
        AddGeneratedInlineFragments(content, element, style, link, source, containingWidth, 0D, 0D, runs,
            insideGeneratedBox: true);
        runs = ApplyScopedFontFallbacks(runs);
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth;
        if (style.ExplicitWidth.HasValue) {
            boxWidth = ResolveBoxWidth(availableWidth, style);
        } else {
            var intrinsic = new List<IntrinsicTextRun> { IntrinsicTextRun.ParagraphStart(style) };
            foreach (HtmlInlineRun run in runs) {
                if (run.AtomicBlock != null) intrinsic.Add(IntrinsicTextRun.Replaced(run.AtomicBlock.Width, run.Style));
                else if (run.Text.Length > 0) intrinsic.Add(new IntrinsicTextRun(run.Text, run.Style));
            }
            IReadOnlyList<IntrinsicTextRun> normalized = NormalizeIntrinsicTextRuns(intrinsic);
            double contentWidth = Math.Min(MeasureMaxContentRuns(normalized),
                Math.Max(MeasureMinContentRuns(normalized), Math.Max(1D, availableWidth - style.HorizontalInsets)));
            // Reuse the normal width constraints after resolving auto to its shrink-to-fit width.
            HtmlRenderBoxStyle sized = style.Clone();
            SetPositionedExplicitWidth(sized, contentWidth + style.HorizontalInsets + style.MarginLeft + style.MarginRight);
            boxWidth = ResolveBoxWidth(availableWidth, sized);
        }

        HtmlInlineLayout inline = LayoutInlineRuns(runs, Math.Max(1D, boxWidth - style.HorizontalInsets), style, element);
        double boxHeight = ResolveBoxHeight(inline.Height, boxWidth, style);
        var visuals = new List<HtmlRenderVisual>();
        AddGeneratedBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element, source);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        var contentVisuals = new List<HtmlRenderVisual>();
        foreach (HtmlRenderVisual visual in inline.Visuals) {
            contentVisuals.Add(visual.Translate(contentX, contentY, contentVisuals.Count));
        }
        // A pseudo box has its own overflow even when its originating element propagates overflow to the viewport.
        AppendOverflowContent(visuals, contentVisuals, style, element,
            style.MarginLeft + style.BorderLeftWidth,
            style.MarginTop + style.BorderTopWidth,
            Math.Max(0.01D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
            Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
            propagateViewportOverflow: false);
        AddGeneratedBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element, source);
        var block = new HtmlRenderFlowBlock(
            style.MarginLeft + boxWidth + style.MarginRight,
            Math.Max(0.01D, style.MarginTop + boxHeight + style.MarginBottom),
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            avoidBreakInside: true,
            source: source);
        // Paint effects need the resolved pseudo width, not the originating element's available width.
        HtmlRenderBoxStyle effectStyle = style.Clone();
        SetPositionedExplicitWidth(effectStyle, block.Width);
        return ApplyPaintEffects(block, effectStyle, containingWidth, source, out _);
    }
}
