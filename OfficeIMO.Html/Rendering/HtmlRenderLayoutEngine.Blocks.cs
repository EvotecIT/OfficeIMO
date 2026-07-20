using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<HtmlRenderFlowBlock> BuildChildBlocks(IElement container, double width, HtmlRenderBoxStyle parentStyle, int depth) =>
        BuildChildBlocks(container, container.ChildNodes, width, parentStyle, depth, includeGeneratedBefore: true, includeGeneratedAfter: true);

    private IReadOnlyList<HtmlRenderFlowBlock> BuildChildBlocks(
        IElement container,
        IEnumerable<INode> nodes,
        double width,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        bool includeGeneratedBefore,
        bool includeGeneratedAfter) {
        EnsureDepth(depth, container);
        var blocks = new List<HtmlRenderFlowBlock>();
        if (includeGeneratedBefore) AddGeneratedContentBlock(blocks, container, HtmlPseudoElementKind.Before, width, parentStyle);
        double flowHeight = blocks.Sum(block => block.Height);
        var adjoiningMargins = new List<double>();
        double allocatedAdjoiningMargins = 0D;
        var inlineNodes = new List<INode>();
        foreach (INode node in nodes) {
            CheckCancellation();
            if (node is IElement element) {
                if (ShouldSkipElement(element)) {
                    continue;
                }

                HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(element, width, parentStyle);
                if (childStyle.Display == "none") {
                    continue;
                }
                if (childStyle.Display == "contents" && HasBlockChildren(element, width, childStyle)) {
                    double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth);
                    flowHeight += inlineHeight;
                    foreach (HtmlRenderFlowBlock flattenedBlock in BuildChildBlocks(element, width, childStyle, depth + 1)) {
                        blocks.Add(flattenedBlock);
                        flowHeight += flattenedBlock.Height;
                    }

                    adjoiningMargins.Clear();
                    allocatedAdjoiningMargins = 0D;
                    continue;
                }
                if (ShouldExtractOutOfFlow(childStyle)) {
                    if (UsesInlineStaticPosition(element, childStyle)) {
                        inlineNodes.Add(node);
                        continue;
                    }
                    double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth);
                    flowHeight += inlineHeight;
                    if (inlineHeight > 0D) {
                        adjoiningMargins.Clear();
                        allocatedAdjoiningMargins = 0D;
                    }
                    PositionedStaticAnchor? staticAnchor = HtmlRenderStyleResolver.IsBlockElement(element, childStyle)
                        ? new PositionedStaticAnchor(container, 0D, flowHeight)
                        : null;
                    RegisterOutOfFlowElement(container, element, childStyle, parentStyle, depth + 1, staticAnchor);
                    continue;
                }

                if (childStyle.FloatSide != "none") {
                    inlineNodes.Add(node);
                    continue;
                }

                if (HtmlRenderStyleResolver.IsBlockElement(element, childStyle)) {
                    double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth);
                    flowHeight += inlineHeight;
                    if (inlineHeight > 0D) {
                        adjoiningMargins.Clear();
                        allocatedAdjoiningMargins = 0D;
                    }
                    HtmlRenderFlowBlock childBlock = LayoutElement(element, width, childStyle, parentStyle, depth + 1);
                    double marginAdjustment = 0D;
                    if (childBlock.HasCollapsibleMargins && adjoiningMargins.Count > 0) {
                        adjoiningMargins.Add(childBlock.CollapsibleMarginTop);
                        allocatedAdjoiningMargins += childBlock.CollapsibleMarginTop;
                        if (!childBlock.CollapsesThrough) {
                            marginAdjustment = allocatedAdjoiningMargins - CollapseVerticalMargins(adjoiningMargins);
                        }
                    }
                    childBlock = childBlock.AdjustLeadingFlowSpace(marginAdjustment);
                    HtmlRenderBoxStyle placementStyle = childStyle.Clone();
                    if (childBlock.HasCollapsibleMargins) {
                        placementStyle.MarginTop = childBlock.CollapsibleMarginTop;
                        placementStyle.MarginBottom = childBlock.CollapsibleMarginBottom;
                    }
                    RecordNormalFlowPlacement(element, container, 0D, flowHeight - marginAdjustment, placementStyle);
                    blocks.Add(childBlock);
                    flowHeight += childBlock.Height;
                    if (!childBlock.HasCollapsibleMargins) {
                        adjoiningMargins.Clear();
                        allocatedAdjoiningMargins = 0D;
                    } else if (childBlock.CollapsesThrough) {
                        if (adjoiningMargins.Count == 0) {
                            adjoiningMargins.Add(childBlock.CollapsibleMarginTop);
                            allocatedAdjoiningMargins = childBlock.CollapsibleMarginTop;
                        }
                        adjoiningMargins.Add(childBlock.CollapsibleMarginBottom);
                        allocatedAdjoiningMargins += childBlock.CollapsibleMarginBottom;
                        double collapsed = CollapseVerticalMargins(adjoiningMargins);
                        double trailingAdjustment = allocatedAdjoiningMargins - collapsed;
                        if (Math.Abs(trailingAdjustment) > 0.0001D) {
                            HtmlRenderFlowBlock adjusted = childBlock.AdjustTrailingFlowSpace(trailingAdjustment);
                            blocks[blocks.Count - 1] = adjusted;
                            flowHeight += adjusted.Height - childBlock.Height;
                            childBlock = adjusted;
                        }
                        allocatedAdjoiningMargins = collapsed;
                    } else {
                        adjoiningMargins.Clear();
                        adjoiningMargins.Add(childBlock.CollapsibleMarginBottom);
                        allocatedAdjoiningMargins = childBlock.CollapsibleMarginBottom;
                    }
                    continue;
                }
            }

            inlineNodes.Add(node);
        }

        double trailingInlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth);
        if (trailingInlineHeight > 0D) adjoiningMargins.Clear();
        if (includeGeneratedAfter) AddGeneratedContentBlock(blocks, container, HtmlPseudoElementKind.After, width, parentStyle);
        return blocks;
    }

    private static double CollapseVerticalMargins(double first, double second) {
        double positive = Math.Max(0D, Math.Max(first, second));
        double negative = Math.Min(0D, Math.Min(first, second));
        return positive + negative;
    }

    private static double CollapseVerticalMargins(IEnumerable<double> margins) {
        double positive = 0D;
        double negative = 0D;
        foreach (double margin in margins) {
            positive = Math.Max(positive, margin);
            negative = Math.Min(negative, margin);
        }
        return positive + negative;
    }

    private HtmlRenderFlowBlock LayoutElement(IElement element, double containingWidth, HtmlRenderBoxStyle style, HtmlRenderBoxStyle parentStyle, int depth) {
        EnsureDepth(depth, element);
        ReportUnsupportedFloatValues(element, style);
        ReportUnsupportedOverflowValues(element, style);
        ReportUnsupportedMultiColumnValues(element, style);
        _layoutStyles[element] = style.Clone();
        string tag = element.TagName.ToLowerInvariant();
        double? containingHeight = ResolveContainingBlockHeight(parentStyle);
        if (tag == "img") return AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(LayoutImage(element, containingWidth, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element);
        if (IsFormControlElement(tag)) return AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(LayoutFormControl(element, containingWidth, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element);
        if (tag == "table") return AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(LayoutTable(element, containingWidth, style, depth), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element);
        if (tag == "hr") return AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(LayoutHorizontalRule(element, containingWidth, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element);
        if (style.Display == "flex" && TryLayoutFlexContainer(element, containingWidth, style, depth, out HtmlRenderFlowBlock flexBlock)) {
            flexBlock = ApplyElementSemantics(flexBlock, element);
            return AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(flexBlock, style, element, containingWidth), style, containingWidth, containingHeight, element), style, element);
        }
        if (style.Display == "grid" && TryLayoutGridContainer(element, containingWidth, style, depth, out HtmlRenderFlowBlock gridBlock)) {
            gridBlock = ApplyElementSemantics(gridBlock, element);
            return AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(gridBlock, style, element, containingWidth), style, containingWidth, containingHeight, element), style, element);
        }
        if (TryLayoutMultiColumnContainer(element, containingWidth, style, depth, out HtmlRenderFlowBlock columnsBlock)) {
            columnsBlock = ApplyElementSemantics(columnsBlock, element);
            return AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(columnsBlock, style, element, containingWidth), style, containingWidth, containingHeight, element), style, element);
        }

        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        var contentVisuals = new List<HtmlRenderVisual>();
        var childPaintLayers = new List<FlowPaintLayer>();
        var contentBreakOffsets = new List<double>();
        var lineBreakOffsets = new List<double>();
        var lineBreakGroups = new List<HtmlRenderLineBreakGroup>();
        var continuationGroups = new List<HtmlRenderContinuationGroup>();
        var trailingGroups = new List<HtmlRenderTrailingGroup>();
        double contentHeight = 0D;
        bool usesBlockFormatting = HasBlockChildren(element, contentWidth, style);
        List<HtmlRenderFlowBlock> children = usesBlockFormatting
            ? BuildChildBlocks(element, contentWidth, style, depth).ToList()
            : new List<HtmlRenderFlowBlock>();

        if (children.Count > 0 && CanCollapseParentMargin(style, top: true) && children[0].HasCollapsibleMargins) {
            HtmlRenderFlowBlock first = children[0];
            double childMargin = first.CollapsibleMarginTop;
            style = style.Clone();
            style.MarginTop = CollapseVerticalMargins(style.MarginTop, childMargin);
            children[0] = first
                .AdjustLeadingFlowSpace(childMargin)
                .WithCollapsibleMargins(0D, first.CollapsibleMarginBottom, first.OwnerElement!);
            if (first.OwnerElement != null) RemoveNormalFlowTopMargin(first.OwnerElement);
        }
        if (children.Count > 0 && CanCollapseParentMargin(style, top: false) && children[children.Count - 1].HasCollapsibleMargins) {
            int lastIndex = children.Count - 1;
            HtmlRenderFlowBlock last = children[lastIndex];
            double childMargin = last.CollapsibleMarginBottom;
            style = style.Clone();
            style.MarginBottom = CollapseVerticalMargins(style.MarginBottom, childMargin);
            children[lastIndex] = last
                .AdjustTrailingFlowSpace(childMargin)
                .WithCollapsibleMargins(last.CollapsibleMarginTop, 0D, last.OwnerElement!);
        }
        _layoutStyles[element] = style.Clone();

        if (usesBlockFormatting) {
            foreach (HtmlRenderFlowBlock child in children) {
                double childStart = contentHeight;
                childPaintLayers.Add(new FlowPaintLayer(child, 0D, childStart, childPaintLayers.Count));

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
            AppendFlowPaintLayers(contentVisuals, childPaintLayers);
        } else {
            string? prefix = tag == "li" ? ResolveListPrefix(element, style) : null;
            HtmlInlineLayout inline = LayoutInlineNodes(element.ChildNodes, contentWidth, style, depth, prefix, element);
            contentVisuals.AddRange(inline.Visuals);
            contentHeight = inline.Height;
            contentBreakOffsets.AddRange(inline.BreakOffsets);
            lineBreakOffsets.AddRange(inline.BreakOffsets);
        }

        if (contentHeight <= 0D && style.ExplicitHeight == null && style.BackgroundColor == null && !style.HasBorderLayout) {
            contentHeight = tag == "div" || tag == "section" || tag == "article" ? 0D : style.LineHeight;
        }

        bool zeroHeightCollapsible = CanUseZeroHeightForMarginCollapse(style, parentStyle, contentHeight);
        double boxHeight = zeroHeightCollapsible ? 0D : ResolveBoxHeight(contentHeight, style);
        double outerHeight = style.MarginTop + boxHeight + style.MarginBottom;
        if (outerHeight <= 0D) outerHeight = 0.01D;
        var visuals = new List<HtmlRenderVisual>();
        var overflowContent = new List<HtmlRenderVisual>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        AppendLocalPositionedVisuals(
            element,
            Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
            Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
            style.MarginLeft + style.BorderLeftWidth,
            style.MarginTop + style.BorderTopWidth,
            PositionedPaintBand.Negative,
            overflowContent);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        foreach (HtmlRenderVisual visual in contentVisuals) {
            overflowContent.Add(visual.Translate(contentX, contentY, overflowContent.Count));
        }
        if (style.Position != "static" || _localPositionedElements.ContainsKey(element)) {
            AppendLocalPositionedVisuals(
                element,
                Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
                Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
                style.MarginLeft + style.BorderLeftWidth,
                style.MarginTop + style.BorderTopWidth,
                PositionedPaintBand.NonNegative,
                overflowContent);
        }
        AppendOverflowContent(
            visuals,
            overflowContent,
            style,
            element,
            style.MarginLeft + style.BorderLeftWidth,
            style.MarginTop + style.BorderTopWidth,
            Math.Max(0.01D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
            Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth));
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);

        ReportUnsupportedLayout(element, style);
        double contentYForBreaks = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
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

        var block = new HtmlRenderFlowBlock(
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
        block = ApplyElementSemantics(block, element);
        bool collapsesThrough = CanCollapseThroughEmptyBlock(style, usesBlockFormatting, children, contentVisuals, contentHeight);
        return AttachElementMargins(ApplyElementPositioning(block, style, containingWidth, containingHeight, element), style, element, collapsesThrough);
    }

    private static HtmlRenderFlowBlock AttachElementMargins(HtmlRenderFlowBlock block, HtmlRenderBoxStyle style, IElement element, bool collapsesThrough = false) =>
        block.WithCollapsibleMargins(style.MarginTop, style.MarginBottom, element, collapsesThrough);

    private static bool CanCollapseParentMargin(HtmlRenderBoxStyle style, bool top) {
        if (style.Display != "block" && style.Display != "list-item") return false;
        if (style.OverflowX != "visible" || style.OverflowY != "visible") return false;
        if (top) return style.BorderTopWidth <= 0D && style.PaddingTop <= 0D;
        return style.BorderBottomWidth <= 0D
            && style.PaddingBottom <= 0D
            && !style.ExplicitHeight.HasValue
            && (!style.MinHeight.HasValue || style.MinHeight.Value <= 0D);
    }

    private static bool CanCollapseThroughEmptyBlock(
        HtmlRenderBoxStyle style,
        bool usesBlockFormatting,
        IReadOnlyList<HtmlRenderFlowBlock> children,
        IReadOnlyList<HtmlRenderVisual> contentVisuals,
        double contentHeight) {
        if (style.Display != "block" || style.OverflowX != "visible" || style.OverflowY != "visible") return false;
        if (style.BorderTopWidth > 0D || style.BorderBottomWidth > 0D || style.PaddingTop > 0D || style.PaddingBottom > 0D) return false;
        if (style.ExplicitHeight.HasValue || style.MinHeight.HasValue && style.MinHeight.Value > 0D) return false;
        if (usesBlockFormatting) return children.Count == 0 || children.All(child => child.CollapsesThrough);
        return contentVisuals.Count == 0 && contentHeight <= 0.0001D;
    }

    private static bool CanUseZeroHeightForMarginCollapse(HtmlRenderBoxStyle style, HtmlRenderBoxStyle parentStyle, double contentHeight) =>
        contentHeight <= 0.0001D
        && style.Display == "block"
        && parentStyle.Display != "flex"
        && parentStyle.Display != "inline-flex"
        && parentStyle.Display != "grid"
        && parentStyle.Display != "inline-grid"
        && style.BorderTopWidth <= 0D
        && style.BorderBottomWidth <= 0D
        && style.PaddingTop <= 0D
        && style.PaddingBottom <= 0D
        && !style.ExplicitHeight.HasValue
        && (!style.MinHeight.HasValue || style.MinHeight.Value <= 0D);

    private double FlushInlineNodes(ICollection<HtmlRenderFlowBlock> blocks, List<INode> nodes, double width, HtmlRenderBoxStyle style, IElement sourceElement, int depth) {
        if (nodes.Count == 0) return 0D;
        HtmlInlineLayout inline = LayoutInlineNodes(nodes, width, style, depth + 1, null, null);
        nodes.Clear();
        if (inline.Height <= 0D || inline.Visuals.Count == 0) return 0D;
        var block = new HtmlRenderFlowBlock(
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
            pageName: style.PageName);
        blocks.Add(block);
        return block.Height;
    }

    private bool HasBlockChildren(IElement element, double width, HtmlRenderBoxStyle parentStyle) {
        foreach (IElement child in element.Children) {
            if (ShouldSkipElement(child)) continue;
            HtmlRenderBoxStyle style = _styleResolver.Resolve(child, width, parentStyle);
            if (style.FloatSide != "none") return true;
            if (style.Display == "contents" && HasBlockChildren(child, width, style)) return true;
            if (style.Display != "none" && ShouldExtractOutOfFlow(style) && !UsesInlineStaticPosition(child, style)) return true;
            if (style.Display != "none" && HtmlRenderStyleResolver.IsBlockElement(child, style)) return true;
            if (style.Display != "none" && ContainsFloatingDescendant(child, width, style)) return true;
        }

        return false;
    }

    private static bool UsesInlineStaticPosition(IElement element, HtmlRenderBoxStyle style) {
        if (style.Display == "inline-flex" || style.Display == "inline-grid" || style.Display == "inline-block") return true;
        if (style.Display != "inline") return false;
        return style.DisplayWasSpecified || !HtmlRenderStyleResolver.IsDefaultBlockElement(element);
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
        IReadOnlyList<HtmlRenderVisual> visuals = style.PaintVisible ? new[] { visual } : Array.Empty<HtmlRenderVisual>();
        return new HtmlRenderFlowBlock(containingWidth, Math.Max(height, 0.01D), visuals, style.BreakBefore, style.BreakAfter, style.AvoidBreakInside, HtmlRenderStyleResolver.DescribeSource(element), pageName: style.PageName);
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

    private void ReportUnsupportedLayout(IElement element, HtmlRenderBoxStyle style) {
        string display = style.Display;
        if (display == "flex" || display == "inline-flex") {
            AddUnsupported(HtmlRenderDiagnosticCodes.FlexLayoutPending, "This flex formatting case is not active yet; children use normal flow.", element);
        } else if (display == "grid" || display == "inline-grid") {
            AddUnsupported(HtmlRenderDiagnosticCodes.GridLayoutPending, "Grid layout is not yet active in the direct HTML renderer; children use normal flow.", element);
        }
    }

    private static bool ShouldSkipElement(IElement element) {
        string tag = element.TagName.ToLowerInvariant();
        if (element.HasAttribute("hidden")) return true;
        if (tag == "input" && string.Equals(element.GetAttribute("type"), "hidden", StringComparison.OrdinalIgnoreCase)) return true;
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

    private static string? ResolveListPrefix(IElement element, HtmlRenderBoxStyle style) {
        if (string.Equals(style.ListStyleType, "none", StringComparison.OrdinalIgnoreCase)) return null;
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
