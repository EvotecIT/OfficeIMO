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
        if (!_generatedContent.TryGetContent(element, kind, out HtmlGeneratedContent content)
            || !_styleResolver.TryResolvePseudo(element, kind, width, parentStyle, out HtmlRenderBoxStyle style)
            || style.Display == "none") {
            return;
        }
        if (CanPositionGeneratedContentLocally(style, parentStyle)) return;

        string source = DescribePseudoSource(element, kind);
        ReportUnsupportedGeneratedLayout(style, source);
        ResolvePositionPaintOffset(style, width, containingHeight, source, out double offsetX, out double offsetY);
        if (content.Fragments.Count == 1
            && content.Fragments[0].Kind == HtmlGeneratedContentFragmentKind.Text
            && content.Fragments[0].Value.Length == 0
            && style.Display is "inline-block" or "inline-flex" or "inline-grid") {
            double boxWidth = ResolveBoxWidth(Math.Max(1D, width - style.MarginLeft - style.MarginRight), style);
            double boxHeight = ResolveBoxHeight(0D, boxWidth, style);
            var visuals = new List<HtmlRenderVisual>();
            AddGeneratedBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element, source);
            AddGeneratedBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element, source);
            var box = new HtmlRenderFlowBlock(
                style.MarginLeft + boxWidth + style.MarginRight,
                Math.Max(0.01D, style.MarginTop + boxHeight + style.MarginBottom),
                visuals,
                style.BreakBefore,
                style.BreakAfter,
                style.AvoidBreakInside,
                source);
            runs.Add(new HtmlInlineRun(box, style, link, source,
                inheritedPaintOffsetX + offsetX, inheritedPaintOffsetY + offsetY, element));
            return;
        }
        AddGeneratedInlineFragments(
            content,
            element,
            style,
            link,
            source,
            width,
            inheritedPaintOffsetX + offsetX,
            inheritedPaintOffsetY + offsetY,
            runs);
    }

    private void AddGeneratedContentBlock(
        ICollection<HtmlRenderFlowBlock> blocks,
        IElement element,
        HtmlPseudoElementKind kind,
        double containingWidth,
        HtmlRenderBoxStyle parentStyle) {
        if (!_generatedContent.TryGetContent(element, kind, out HtmlGeneratedContent content)
            || !_styleResolver.TryResolvePseudo(element, kind, containingWidth, parentStyle, out HtmlRenderBoxStyle style)
            || style.Display == "none") {
            return;
        }
        if (CanPositionGeneratedContentLocally(style, parentStyle)) return;

        blocks.Add(CreateGeneratedContentBlock(element, kind, containingWidth, parentStyle, style, content));
    }

    private HtmlRenderFlowBlock CreateGeneratedContentBlock(
        IElement element,
        HtmlPseudoElementKind kind,
        double containingWidth,
        HtmlRenderBoxStyle parentStyle,
        HtmlRenderBoxStyle style,
        HtmlGeneratedContent content,
        bool positioned = false) {

        string source = DescribePseudoSource(element, kind);
        ReportUnsupportedGeneratedLayout(style, source);
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        string? link = string.Equals(element.TagName, "a", StringComparison.OrdinalIgnoreCase)
            ? ResolveSafeLink(element.GetAttribute("href"), element)
            : null;
        var runs = new List<HtmlInlineRun>();
        AddGeneratedInlineFragments(content, element, style, link, source, contentWidth, 0D, 0D, runs);
        HtmlInlineLayout inline = LayoutInlineRuns(runs, contentWidth, style);
        double boxHeight = ResolveBoxHeight(inline.Height, boxWidth, style);
        double outerHeight = Math.Max(0.01D, style.MarginTop + boxHeight + style.MarginBottom);
        var visuals = new List<HtmlRenderVisual>();
        bool paintsBlockBox = style.Display == "block" || style.Display == "flow-root" || style.Display == "list-item";
        if (paintsBlockBox) AddGeneratedBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element, source);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        foreach (HtmlRenderVisual visual in inline.Visuals) {
            visuals.Add(visual.Translate(contentX, contentY, visuals.Count));
        }
        if (paintsBlockBox) AddGeneratedBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element, source);

        IEnumerable<double> breakOffsets = inline.BreakOffsets
            .Select(offset => contentY + offset)
            .Concat(new[] { outerHeight });
        var block = new HtmlRenderFlowBlock(
            positioned ? style.MarginLeft + boxWidth + style.MarginRight : containingWidth,
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
        return positioned
            ? block
            : ApplyPositioning(block, style, containingWidth, ResolveContainingBlockHeight(parentStyle), source);
    }

    private static bool CanPositionGeneratedContentLocally(HtmlRenderBoxStyle style, HtmlRenderBoxStyle parentStyle) =>
        style.Position == "absolute"
        && (style.Display is "block" or "flow-root" or "list-item")
        && style.ExplicitWidth.HasValue
        && style.ExplicitHeight.HasValue
        && (parentStyle.Position is "relative" or "sticky")
        && (parentStyle.Display is "block" or "flow-root" or "list-item" or "inline-block")
        && (!IsAutoInset(style.Left) || !IsAutoInset(style.Right))
        && (!IsAutoInset(style.Top) || !IsAutoInset(style.Bottom));

    private bool TryGetLocallyPositionedGeneratedContentZIndex(
        IElement element,
        HtmlPseudoElementKind kind,
        double containingWidth,
        HtmlRenderBoxStyle parentStyle,
        out int zIndex) {
        zIndex = 0;
        if (!_generatedContent.TryGetContent(element, kind, out _)
            || !_styleResolver.TryResolvePseudo(element, kind, containingWidth, parentStyle, out HtmlRenderBoxStyle style)
            || !CanPositionGeneratedContentLocally(style, parentStyle)) return false;
        zIndex = ResolvePositionedZIndex(element, style);
        return true;
    }

    private void AppendLocallyPositionedGeneratedContent(
        IElement element,
        HtmlPseudoElementKind kind,
        double containingWidth,
        double containingHeight,
        double originX,
        double originY,
        HtmlRenderBoxStyle parentStyle,
        ICollection<HtmlRenderVisual> visuals) {
        if (!_generatedContent.TryGetContent(element, kind, out HtmlGeneratedContent content)
            || !_styleResolver.TryResolvePseudo(element, kind, containingWidth, parentStyle, out HtmlRenderBoxStyle style)
            || !CanPositionGeneratedContentLocally(style, parentStyle)) return;

        string source = DescribePseudoSource(element, kind);
        HtmlRenderFlowBlock block = CreateGeneratedContentBlock(
            element, kind, containingWidth, parentStyle, style, content, positioned: true);
        block = ApplyPaintEffects(block, style, containingWidth, source, out _);
        double? left = ResolveOutOfFlowInset(style.Left, containingWidth, style, source, "left");
        double? right = ResolveOutOfFlowInset(style.Right, containingWidth, style, source, "right");
        double? top = ResolveOutOfFlowInset(style.Top, containingHeight, style, source, "top");
        double? bottom = ResolveOutOfFlowInset(style.Bottom, containingHeight, style, source, "bottom");
        double x = left ?? containingWidth - (right ?? 0D) - block.Width;
        double y = top ?? containingHeight - (bottom ?? 0D) - block.Height;
        foreach (HtmlRenderVisual visual in block.Visuals) {
            visuals.Add(visual.Translate(originX + x, originY + y, visuals.Count));
        }
    }

    private void AddGeneratedInlineFragments(
        HtmlGeneratedContent content,
        IElement element,
        HtmlRenderBoxStyle style,
        string? link,
        string source,
        double containingWidth,
        double paintOffsetX,
        double paintOffsetY,
        ICollection<HtmlInlineRun> runs) {
        for (int index = 0; index < content.Fragments.Count; index++) {
            HtmlGeneratedContentFragment fragment = content.Fragments[index];
            string fragmentSource = fragment.Kind == HtmlGeneratedContentFragmentKind.Text
                ? source
                : source + ":content-" + fragment.Kind.ToString().ToLowerInvariant()
                    + "[" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
            if (fragment.Kind == HtmlGeneratedContentFragmentKind.Text) {
                string text = ApplyTextTransform(fragment.Value, style);
                if (text.Length > 0 && style.Font.Size > 0D) {
                    runs.Add(new HtmlInlineRun(text, style, link, fragmentSource, paintOffsetX, paintOffsetY, element));
                }
                continue;
            }

            if (fragment.Kind == HtmlGeneratedContentFragmentKind.Leader) {
                runs.Add(new HtmlInlineRun(
                    string.Empty,
                    style,
                    link,
                    fragmentSource,
                    paintOffsetX,
                    paintOffsetY,
                    element,
                    logicalText: string.Empty,
                    leaderPattern: fragment.Value));
                continue;
            }

            if (fragment.Kind == HtmlGeneratedContentFragmentKind.TargetPage) {
                int pageNumber = _generatedContent.TryGetTargetPage(fragment.Value, out int resolvedPage) ? resolvedPage : 8888;
                string counterStyle = string.IsNullOrWhiteSpace(fragment.Format) ? "decimal" : fragment.Format!;
                string pageText = _counterStyles.TryFormat(pageNumber, counterStyle, out string custom, out _)
                    ? custom
                    : HtmlCounterStyleFormatter.TryFormat(pageNumber, counterStyle, out string standard, out _)
                        ? standard
                        : pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
                runs.Add(new HtmlInlineRun(pageText, style, link, fragmentSource, paintOffsetX, paintOffsetY, element));
                continue;
            }

            IDocument? owner = element.Owner;
            if (owner == null) continue;
            IElement imageElement = owner.CreateElement("img");
            imageElement.SetAttribute("src", fragment.Value);
            double imageWidth = ResolveFloatingImageOuterWidth(imageElement, style);
            HtmlRenderFlowBlock image = LayoutImage(imageElement, imageWidth, style, link);
            runs.Add(new HtmlInlineRun(
                image,
                style,
                link,
                fragmentSource,
                paintOffsetX,
                paintOffsetY,
                element,
                isReplacedImage: true));
        }
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
        if (!style.PaintVisible) return;
        HtmlResolvedBorderRadii radii = ResolveBoxRadii(style, width, height, element, source);
        AddOuterBoxShadows(visuals, style, x, y, width, height, radii, element, source);
        AddBoxBackgroundCore(visuals, style, x, y, width, height, style.BorderInsets, radii, element, source, source);
        AddInsetBoxShadows(visuals, style, x, y, width, height, radii, element, source);
        AddBorderPaint(visuals, style, x, y, width, height, radii, element, source);
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
        if (!style.PaintVisible) return;
        HtmlResolvedBorderRadii radii = ResolveBoxRadii(style, width, height, element, source);
        AddOutlinePaint(visuals, style, x, y, width, height, radii, element, source);
    }

    private static string DescribePseudoSource(IElement element, HtmlPseudoElementKind kind) =>
        HtmlRenderStyleResolver.DescribeSource(element)
        + (kind switch {
            HtmlPseudoElementKind.Before => "::before",
            HtmlPseudoElementKind.After => "::after",
            HtmlPseudoElementKind.Marker => "::marker",
            HtmlPseudoElementKind.FootnoteCall => "::footnote-call",
            _ => "::footnote-marker"
        });

    private void ReportUnsupportedGeneratedLayout(HtmlRenderBoxStyle style, string source) {
        string display = style.Display;
        if (display == "flex" || display == "inline-flex") {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.FlexLayoutPending, "Flex layout is not yet active for generated content; text uses normal flow.", HtmlDiagnosticSeverity.Warning, source);
        } else if (display == "grid" || display == "inline-grid") {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.GridLayoutPending, "Grid layout is not yet active for generated content; text uses normal flow.", HtmlDiagnosticSeverity.Warning, source);
        }
    }
}
