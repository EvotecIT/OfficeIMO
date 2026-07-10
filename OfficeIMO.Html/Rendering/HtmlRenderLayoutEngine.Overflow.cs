using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock ApplyOverflowToSpecializedBlock(
        HtmlRenderFlowBlock block,
        HtmlRenderBoxStyle style,
        IElement element,
        double containingWidth) {
        if (style.OverflowX == "visible" && style.OverflowY == "visible") return block;
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double boxHeight = Math.Max(0.01D, block.Height - style.MarginTop - style.MarginBottom);
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        var outside = new List<HtmlRenderVisual>();
        var content = new List<HtmlRenderVisual>();
        foreach (HtmlRenderVisual visual in block.Visuals) {
            if (visual is HtmlRenderShape
                && string.Equals(visual.Source, source, StringComparison.Ordinal)
                && Math.Abs(visual.X - style.MarginLeft) <= 0.0001D
                && Math.Abs(visual.Y - style.MarginTop) <= 0.0001D
                && Math.Abs(visual.Width - boxWidth) <= 0.0001D
                && Math.Abs(visual.Height - boxHeight) <= 0.0001D) {
                outside.Add(visual.Translate(0D, 0D, outside.Count));
            } else {
                content.Add(visual.Translate(0D, 0D, content.Count));
            }
        }
        AppendOverflowContent(
            outside,
            content,
            style,
            element,
            style.MarginLeft + style.BorderWidth,
            style.MarginTop + style.BorderWidth,
            Math.Max(0.01D, boxWidth - style.BorderWidth * 2D),
            Math.Max(0.01D, boxHeight - style.BorderWidth * 2D));
        return block.WithVisuals(outside);
    }

    private void ReportUnsupportedOverflowValues(IElement element, HtmlRenderBoxStyle style) {
        if (style.UnsupportedOverflowX.Length == 0 && style.UnsupportedOverflowY.Length == 0) return;
        if (!_reportedOverflowValueFallbacks.Add(element)) return;
        var details = new List<string>(2);
        if (style.UnsupportedOverflowX.Length > 0) details.Add("overflow-x=" + style.UnsupportedOverflowX);
        if (style.UnsupportedOverflowY.Length > 0) details.Add("overflow-y=" + style.UnsupportedOverflowY);
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.OverflowValueUnsupported,
            "An overflow value used the visible fallback.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            string.Join(";", details.Distinct(StringComparer.Ordinal)));
    }

    private void AppendOverflowContent(
        ICollection<HtmlRenderVisual> target,
        IReadOnlyList<HtmlRenderVisual> content,
        HtmlRenderBoxStyle style,
        IElement element,
        double clipX,
        double clipY,
        double clipWidth,
        double clipHeight) {
        bool clipHorizontal = style.OverflowX != "visible";
        bool clipVertical = style.OverflowY != "visible";
        if (!clipHorizontal && !clipVertical) {
            foreach (HtmlRenderVisual visual in content) target.Add(visual.Translate(0D, 0D, target.Count));
            return;
        }

        if (ShouldReportOverflowSnapshot(content, style, clipX, clipY, clipWidth, clipHeight)
            && _reportedOverflowScrollSnapshots.Add(element)) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.OverflowScrollSnapshot,
                "A scrollable overflow box was clipped at its initial static scroll position without interactive scrollbars.",
                HtmlDiagnosticSeverity.Info,
                HtmlRenderStyleResolver.DescribeSource(element),
                "overflow-x=" + style.OverflowX + ";overflow-y=" + style.OverflowY);
        }

        if (content.Count == 0) return;
        target.Add(new HtmlRenderClipGroup(
            clipX,
            clipY,
            Math.Max(0.01D, clipWidth),
            Math.Max(0.01D, clipHeight),
            clipHorizontal,
            clipVertical,
            content,
            target.Count,
            HtmlRenderStyleResolver.DescribeSource(element)));
    }

    private static bool ShouldReportOverflowSnapshot(
        IReadOnlyList<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double clipX,
        double clipY,
        double clipWidth,
        double clipHeight) {
        bool horizontalScroll = style.OverflowX == "scroll";
        bool verticalScroll = style.OverflowY == "scroll";
        double right = clipX + clipWidth;
        double bottom = clipY + clipHeight;
        foreach (HtmlRenderVisual visual in visuals) {
            if (style.OverflowX == "auto" && (visual.X < clipX - 0.0001D || visual.X + visual.Width > right + 0.0001D)) horizontalScroll = true;
            if (style.OverflowY == "auto" && (visual.Y < clipY - 0.0001D || visual.Y + visual.Height > bottom + 0.0001D)) verticalScroll = true;
            if (horizontalScroll && verticalScroll) break;
        }
        return horizontalScroll || verticalScroll;
    }
}
