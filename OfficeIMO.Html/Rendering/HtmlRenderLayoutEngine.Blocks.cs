using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<HtmlRenderFlowBlock> BuildChildBlocks(IElement container, double width, HtmlRenderBoxStyle parentStyle, int depth) {
        EnsureDepth(depth, container);
        var blocks = new List<HtmlRenderFlowBlock>();
        var inlineNodes = new List<INode>();
        foreach (INode node in container.ChildNodes) {
            if (node is IElement element) {
                if (ShouldSkipElement(element)) {
                    continue;
                }

                HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(element, width, parentStyle);
                if (childStyle.Display == "none") {
                    continue;
                }

                if (HtmlRenderStyleResolver.IsBlockElement(element, childStyle)) {
                    FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth);
                    blocks.Add(LayoutElement(element, width, childStyle, parentStyle, depth + 1));
                    continue;
                }
            }

            inlineNodes.Add(node);
        }

        FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth);
        return blocks;
    }

    private HtmlRenderFlowBlock LayoutElement(IElement element, double containingWidth, HtmlRenderBoxStyle style, HtmlRenderBoxStyle parentStyle, int depth) {
        EnsureDepth(depth, element);
        string tag = element.TagName.ToLowerInvariant();
        if (tag == "img") return LayoutImage(element, containingWidth, style);
        if (tag == "table") return LayoutTable(element, containingWidth, style, depth);
        if (tag == "hr") return LayoutHorizontalRule(element, containingWidth, style);

        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        var contentVisuals = new List<HtmlRenderVisual>();
        var contentBreakOffsets = new List<double>();
        var lineBreakOffsets = new List<double>();
        var lineBreakGroups = new List<HtmlRenderLineBreakGroup>();
        var continuationGroups = new List<HtmlRenderContinuationGroup>();
        var trailingGroups = new List<HtmlRenderTrailingGroup>();
        double contentHeight = 0D;
        IReadOnlyList<HtmlRenderFlowBlock> children = HasBlockChildren(element, contentWidth, style)
            ? BuildChildBlocks(element, contentWidth, style, depth)
            : Array.Empty<HtmlRenderFlowBlock>();

        if (children.Count > 0) {
            foreach (HtmlRenderFlowBlock child in children) {
                double childStart = contentHeight;
                foreach (HtmlRenderVisual visual in child.Visuals) {
                    contentVisuals.Add(visual.Translate(0D, childStart, contentVisuals.Count));
                }

                contentHeight += child.Height;
                foreach (double offset in child.BreakOffsets) {
                    contentBreakOffsets.Add(childStart + offset);
                }

                foreach (HtmlRenderLineBreakGroup group in child.LineBreakGroups) {
                    lineBreakGroups.Add(group.Translate(childStart));
                }

                foreach (HtmlRenderContinuationGroup group in child.ContinuationGroups) {
                    continuationGroups.Add(group.Translate(0D, childStart));
                }

                foreach (HtmlRenderTrailingGroup group in child.TrailingGroups) {
                    trailingGroups.Add(group.Translate(0D, childStart));
                }

                contentBreakOffsets.Add(contentHeight);
            }
        } else {
            string? prefix = tag == "li" ? ResolveListPrefix(element) : null;
            HtmlInlineLayout inline = LayoutInlineNodes(element.ChildNodes, contentWidth, style, depth, prefix);
            contentVisuals.AddRange(inline.Visuals);
            contentHeight = inline.Height;
            contentBreakOffsets.AddRange(inline.BreakOffsets);
            lineBreakOffsets.AddRange(inline.BreakOffsets);
        }

        if (contentHeight <= 0D && style.ExplicitHeight == null && style.BackgroundColor == null && style.BorderWidth <= 0D) {
            contentHeight = tag == "div" || tag == "section" || tag == "article" ? 0D : style.LineHeight;
        }

        double boxHeight = ResolveBoxHeight(contentHeight, style);
        double outerHeight = style.MarginTop + boxHeight + style.MarginBottom;
        if (outerHeight <= 0D) outerHeight = 0.01D;
        var visuals = new List<HtmlRenderVisual>();
        AddBoxShape(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        double contentX = style.MarginLeft + style.BorderWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderWidth + style.PaddingTop;
        foreach (HtmlRenderVisual visual in contentVisuals) {
            visuals.Add(visual.Translate(contentX, contentY, visuals.Count));
        }

        ReportUnsupportedLayout(element, style);
        double contentYForBreaks = style.MarginTop + style.BorderWidth + style.PaddingTop;
        IEnumerable<double> breakOffsets = contentBreakOffsets.Select(offset => contentYForBreaks + offset)
            .Concat(new[] { outerHeight });
        IEnumerable<double> adjustedLineBreakOffsets = lineBreakOffsets.Select(offset => contentYForBreaks + offset);
        IEnumerable<HtmlRenderLineBreakGroup> adjustedLineBreakGroups = lineBreakGroups.Select(group => group.Translate(contentYForBreaks));
        IEnumerable<HtmlRenderContinuationGroup> adjustedContinuationGroups = continuationGroups.Select(group => group.Translate(contentX, contentYForBreaks));
        IEnumerable<HtmlRenderTrailingGroup> adjustedTrailingGroups = trailingGroups.Select(group =>
            group.Translate(
                contentX,
                contentYForBreaks,
                group.SourceEndsAt >= contentHeight - 0.0001D ? outerHeight : (double?)null));
        string? pageName = style.PageName;
        if (pageName == null && children.Count > 0 && children.All(child => string.Equals(child.PageName, children[0].PageName, StringComparison.OrdinalIgnoreCase))) {
            pageName = children[0].PageName;
        }

        return new HtmlRenderFlowBlock(
            containingWidth,
            outerHeight,
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            style.AvoidBreakInside,
            HtmlRenderStyleResolver.DescribeSource(element),
            breakOffsets,
            adjustedLineBreakOffsets,
            style.Orphans,
            style.Widows,
            adjustedLineBreakGroups,
            adjustedContinuationGroups,
            adjustedTrailingGroups,
            pageName: pageName);
    }

    private void FlushInlineNodes(ICollection<HtmlRenderFlowBlock> blocks, List<INode> nodes, double width, HtmlRenderBoxStyle style, IElement sourceElement, int depth) {
        if (nodes.Count == 0) return;
        HtmlInlineLayout inline = LayoutInlineNodes(nodes, width, style, depth + 1, null);
        nodes.Clear();
        if (inline.Height <= 0D || inline.Visuals.Count == 0) return;
        blocks.Add(new HtmlRenderFlowBlock(
            width,
            inline.Height,
            inline.Visuals,
            HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None,
            false,
            HtmlRenderStyleResolver.DescribeSource(sourceElement),
            inline.BreakOffsets,
            inline.BreakOffsets,
            style.Orphans,
            style.Widows,
            pageName: style.PageName));
    }

    private bool HasBlockChildren(IElement element, double width, HtmlRenderBoxStyle parentStyle) {
        foreach (IElement child in element.Children) {
            if (ShouldSkipElement(child)) continue;
            HtmlRenderBoxStyle style = _styleResolver.Resolve(child, width, parentStyle);
            if (style.Display != "none" && HtmlRenderStyleResolver.IsBlockElement(child, style)) return true;
        }

        return false;
    }

    private HtmlRenderFlowBlock LayoutHorizontalRule(IElement element, double containingWidth, HtmlRenderBoxStyle style) {
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double width = ResolveBoxWidth(availableWidth, style);
        double lineWidth = style.BorderWidth > 0D ? style.BorderWidth : 1D;
        var shape = OfficeShape.Rectangle(width, lineWidth);
        shape.FillColor = style.BorderColor;
        shape.StrokeWidth = 0D;
        double height = style.MarginTop + lineWidth + style.MarginBottom;
        var visual = new HtmlRenderShape(shape, style.MarginLeft, style.MarginTop, 0, source: HtmlRenderStyleResolver.DescribeSource(element));
        return new HtmlRenderFlowBlock(containingWidth, Math.Max(height, 0.01D), new[] { visual }, style.BreakBefore, style.BreakAfter, style.AvoidBreakInside, HtmlRenderStyleResolver.DescribeSource(element), pageName: style.PageName);
    }

    private double ResolveBoxWidth(double availableWidth, HtmlRenderBoxStyle style) {
        double width = style.ExplicitWidth ?? (style.BorderBox ? availableWidth : Math.Max(1D, availableWidth - style.HorizontalInsets));
        if (!style.BorderBox) width += style.HorizontalInsets;
        if (style.MinWidth.HasValue) width = Math.Max(width, style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        if (style.MaxWidth.HasValue) width = Math.Min(width, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        return Math.Max(1D, Math.Min(width, availableWidth));
    }

    private static double ResolveBoxHeight(double contentHeight, HtmlRenderBoxStyle style) {
        double height = style.ExplicitHeight ?? contentHeight;
        if (!style.BorderBox || !style.ExplicitHeight.HasValue) height += style.VerticalInsets;
        if (style.MinHeight.HasValue) height = Math.Max(height, style.MinHeight.Value + (style.BorderBox ? 0D : style.VerticalInsets));
        if (style.MaxHeight.HasValue) height = Math.Min(height, style.MaxHeight.Value + (style.BorderBox ? 0D : style.VerticalInsets));
        return Math.Max(0.01D, height);
    }

    private static void AddBoxShape(ICollection<HtmlRenderVisual> visuals, HtmlRenderBoxStyle style, double x, double y, double width, double height, IElement source) {
        if ((!style.BackgroundColor.HasValue || style.BackgroundColor.Value.A == 0) && style.BorderWidth <= 0D) return;
        OfficeShape shape = OfficeShape.Rectangle(width, height);
        shape.FillColor = style.BackgroundColor;
        shape.StrokeColor = style.BorderWidth > 0D ? style.BorderColor : null;
        shape.StrokeWidth = style.BorderWidth;
        visuals.Add(new HtmlRenderShape(shape, x, y, visuals.Count, source: HtmlRenderStyleResolver.DescribeSource(source)));
    }

    private void ReportUnsupportedLayout(IElement element, HtmlRenderBoxStyle style) {
        string display = style.Display;
        if (display == "flex" || display == "inline-flex") {
            AddUnsupported(HtmlRenderDiagnosticCodes.FlexLayoutPending, "Flex layout is not yet active in the direct HTML renderer; children use normal flow.", element);
        } else if (display == "grid" || display == "inline-grid") {
            AddUnsupported(HtmlRenderDiagnosticCodes.GridLayoutPending, "Grid layout is not yet active in the direct HTML renderer; children use normal flow.", element);
        }
    }

    private static bool ShouldSkipElement(IElement element) {
        string tag = element.TagName.ToLowerInvariant();
        return tag == "head" || tag == "style" || tag == "script" || tag == "template" || tag == "noscript" || tag == "meta" || tag == "link" || tag == "title" || tag == "base";
    }

    private void EnsureDepth(int depth, IElement element) {
        if (depth <= _options.MaxLayoutDepth) return;
        throw new HtmlDomLimitException(
            HtmlRenderDiagnosticCodes.DepthLimitExceeded,
            "HTML layout exceeded the configured maximum depth at " + HtmlRenderStyleResolver.DescribeSource(element) + ".",
            nameof(HtmlRenderOptions.MaxLayoutDepth),
            depth,
            _options.MaxLayoutDepth);
    }

    private static string? ResolveListPrefix(IElement element) {
        IElement? parent = element.ParentElement;
        if (parent == null) return "• ";
        if (!string.Equals(parent.TagName, "ol", StringComparison.OrdinalIgnoreCase)) return "• ";
        int start = 1;
        int.TryParse(parent.GetAttribute("start"), out start);
        if (start == 0) start = 1;
        int index = 0;
        foreach (IElement sibling in parent.Children) {
            if (!string.Equals(sibling.TagName, "li", StringComparison.OrdinalIgnoreCase)) continue;
            if (ReferenceEquals(sibling, element)) break;
            index++;
        }

        return (start + index).ToString(System.Globalization.CultureInfo.InvariantCulture) + ". ";
    }
}
