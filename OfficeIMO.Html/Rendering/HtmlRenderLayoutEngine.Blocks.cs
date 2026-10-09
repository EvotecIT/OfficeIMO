using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void ChargeLayoutOperation(string source) {
        ChargeLayoutOperations(1L, source);
    }

    private void ChargeLayoutOperations(long count, string source) {
        _operationBudget.ChargeLayoutOperations(count, _options.MaxLayoutOperations, source);
    }

    private HtmlRenderFlowBlock LayoutElement(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        IElement? continuationTarget = null,
        int continuationLogicalCharacters = 0,
        PagedFloatBoundary? pageBoundary = null,
        IReadOnlyList<HtmlFloatExclusion>? inheritedFloats = null,
        ICollection<HtmlFloatExclusion>? emittedFloats = null, InlineFloatContext? surroundingFloats = null) {
        IElement? root = _document.Body ?? _document.DocumentElement;
        bool tracksPageViewport = _options.Mode == HtmlRenderMode.Paged
            && depth == 1
            && ReferenceEquals(element.ParentElement, root);
        HtmlRenderFlowBlock StampViewport(HtmlRenderFlowBlock block) =>
            tracksPageViewport
                ? block.WithLayoutViewport(_activePageGeometry.Width, _activePageGeometry.Height)
                : block;
        EnsureDepth(depth, element);
        ChargeLayoutOperation(HtmlRenderStyleResolver.DescribeSource(element));
        bool continuesThisBox = continuationTarget != null
            && ContainsElementOrSelf(element, continuationTarget)
            && (!ReferenceEquals(element, continuationTarget) || continuationLogicalCharacters > 0);
        // Page and continuation reflow can enter here with a fresh style,
        // bypassing normal child placement and its auto-margin preparation.
        if (style.HasIntrinsicWidths && !style.IntrinsicWidthsResolved) {
            style = ResolveOrdinaryIntrinsicWidths(element, style, containingWidth, depth);
            style = ResolveNormalFlowHorizontalAutoMargins(element, style, containingWidth);
        }
        style = PrepareButtonChildStyle(element, style);
        if (UsesButtonChildLayout(element) && !style.ExplicitWidth.HasValue) {
            // A native button remains intrinsically sized with display:block.
            // Its descendants still use normal layout, including styled boxes.
            SetPositionedExplicitWidth(style,
                ResolvePositionedOuterWidth(element, style, containingWidth, null, null, depth));
        }
        if (continuesThisBox) style = SuppressContinuationStartDecorations(style);
        _inlineFloatOverhangs.Remove(element);
        ReportUnsupportedFloatValues(element, style);
        ReportUnsupportedOverflowValues(element, style);
        ReportUnsupportedMultiColumnValues(element, style);
        _layoutStyles[element] = style.Clone();
        string tag = element.TagName.ToLowerInvariant();
        double? containingHeight = ResolveContainingBlockHeight(parentStyle);
        if (IsReplacedImageElementTag(tag)) return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(ApplySpecializedElementSemantics(LayoutImage(element, containingWidth, style), element, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        if (tag == "math" && TryLayoutMath(element, containingWidth, style, inheritedLink: null, shrinkToFit: false, out HtmlRenderFlowBlock mathBlock, out _)) return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(ApplySpecializedElementSemantics(mathBlock, element, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        if (IsFormControlElement(tag) && !UsesButtonChildLayout(element)) return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(ApplySpecializedElementSemantics(LayoutFormControl(element, containingWidth, style), element, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        if (tag == "table") return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(ApplySpecializedElementSemantics(LayoutTable(element, containingWidth, style, depth, continuationTarget), element, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        if (tag == "hr") return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(ApplySpecializedElementSemantics(LayoutHorizontalRule(element, containingWidth, style), element, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        if (style.Display == "flex" && TryLayoutFlexContainer(element, containingWidth, style, depth, continuationTarget, pageBoundary, out HtmlRenderFlowBlock flexBlock)) {
            flexBlock = ApplyElementSemantics(flexBlock, element, style);
            return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(flexBlock, style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        }
        if (style.Display == "grid" && TryLayoutGridContainer(element, containingWidth, style, depth, out HtmlRenderFlowBlock gridBlock)) {
            gridBlock = ApplyElementSemantics(gridBlock, element, style);
            return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(gridBlock, style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        }
        if (TryLayoutMultiColumnContainer(element, containingWidth, style, depth, out HtmlRenderFlowBlock columnsBlock)) {
            columnsBlock = ApplyElementSemantics(columnsBlock, element, style);
            return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(columnsBlock, style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        }

        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        bool ownsFloatContext = surroundingFloats == null || EstablishesIndependentFloatContext(style);
        InlineFloatContext contentFloats = ownsFloatContext
            ? new InlineFloatContext(contentWidth, inheritedFloats)
            : surroundingFloats!.At(contentWidth, style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft,
                style.MarginTop + style.BorderTopWidth + style.PaddingTop);
        double positionedContainingWidth = Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth);
        bool hasLocallyPositionedPseudo =
            TryGetLocallyPositionedGeneratedContentZIndex(element, HtmlPseudoElementKind.Before, positionedContainingWidth, style, out _)
            || TryGetLocallyPositionedGeneratedContentZIndex(element, HtmlPseudoElementKind.After, positionedContainingWidth, style, out _);
        var contentVisuals = new List<HtmlRenderVisual>();
        var childPaintLayers = new List<FlowPaintLayer>();
        var contentBreakOffsets = new List<double>();
        var contentAvoidBreakRanges = new List<HtmlRenderAvoidBreakRange>();
        var forcedBreaks = new List<HtmlRenderForcedBreak>();
        var lineBreakOffsets = new List<double>();
        var lineBreakGroups = new List<HtmlRenderLineBreakGroup>();
        var continuationGroups = new List<HtmlRenderContinuationGroup>();
        var trailingGroups = new List<HtmlRenderTrailingGroup>();
        var runningStringAssignments = new List<HtmlCssRunningStringAssignment>();
        var continuationBreakProgress = new List<HtmlInlineBreakProgress>();
        HtmlInlineLayout? inlineLayout = null;
        double contentHeight = 0D;
        bool usesBlockFormatting = HasBlockChildren(element, contentWidth, style, depth);
        IElement? descendantContinuationTarget = continuationTarget != null && !ReferenceEquals(element, continuationTarget)
            ? continuationTarget
            : null;
        bool usesVerticalBlockFormatting = usesBlockFormatting && IsVerticalWritingMode(style.WritingMode);
        double childContainingWidth = usesVerticalBlockFormatting
            ? ResolveVerticalInlineExtent(style, parentStyle, contentWidth)
            : contentWidth;
        List<HtmlFloatExclusion>? childContentFloats = emittedFloats != null ? new List<HtmlFloatExclusion>() : null;
        List<HtmlRenderFlowBlock> children = usesBlockFormatting
            ? BuildChildBlocks(
                element,
                childContainingWidth,
                style,
                depth,
                descendantContinuationTarget,
                descendantContinuationTarget == null ? 0 : continuationLogicalCharacters,
                pageBoundary?.Shift(style.MarginTop + style.BorderTopWidth + style.PaddingTop),
                inheritedFloats,
                EstablishesFloatContainingBlock(style) ? null : childContentFloats, contentFloats).ToList()
            : new List<HtmlRenderFlowBlock>();
        if (usesBlockFormatting && _inlineFloatOverhangs.ContainsKey(element)
            && (children.Count != 1 || children[0].OwnerElement != null)) {
            // An earlier inline run does not establish a float overhang for the
            // element when subsequent block content follows it.
            _inlineFloatOverhangs.Remove(element);
        }

        if (!usesBlockFormatting || !usesVerticalBlockFormatting && children.Count == 0) {
            HtmlListMarker? marker = tag == "li" ? ResolveListMarker(element, style, contentWidth) : null;
            int skipped = ReferenceEquals(element, continuationTarget) ? continuationLogicalCharacters : 0;
            double extent = IsVerticalWritingMode(style.WritingMode)
                ? ResolveVerticalInlineExtent(style, parentStyle, contentWidth) : contentWidth;
            inlineLayout = LayoutInlineNodes(element.ChildNodes, extent, style, depth, marker, element, skipped, pageBoundary?.Shift(style.MarginTop + style.BorderTopWidth + style.PaddingTop), inheritedFloats, floatContext: contentFloats);
            if (inlineLayout.InterruptedFlow != null) {
                children = new List<HtmlRenderFlowBlock> { inlineLayout.InterruptedFlow };
                inlineLayout = null;
                usesBlockFormatting = true;
                usesVerticalBlockFormatting = IsVerticalWritingMode(style.WritingMode);
            }
        }

        double collapsedChildTopMargin = 0D;
        HtmlCollapsedMargin collapsedTopGroup = new HtmlCollapsedMargin(style.MarginTop);
        HtmlCollapsedMargin collapsedBottomGroup = new HtmlCollapsedMargin(style.MarginBottom);
        if (!usesVerticalBlockFormatting && children.Count > 0 && CanCollapseParentMargin(style, top: true) && children[0].HasCollapsibleMargins) {
            HtmlRenderFlowBlock first = children[0];
            double childMargin = first.CollapsibleMarginTop;
            collapsedChildTopMargin = childMargin;
            style = style.Clone();
            collapsedTopGroup = collapsedTopGroup.Combine(first.CollapsibleMarginTopGroup);
            style.MarginTop = collapsedTopGroup.Value;
            children[0] = first
                .AdjustLeadingFlowSpace(childMargin)
                .WithCollapsibleMargins(0D, first.CollapsibleMarginBottom, first.OwnerElement!,
                    bottomGroup: first.CollapsibleMarginBottomGroup);
            if (first.OwnerElement != null) RemoveNormalFlowTopMargin(first.OwnerElement);
        }
        if (emittedFloats != null && childContentFloats != null) {
            foreach (HtmlFloatExclusion exclusion in childContentFloats) {
                emittedFloats.Add(exclusion.Shift(0D, -collapsedChildTopMargin));
            }
        }
        if (!usesVerticalBlockFormatting && children.Count > 0 && CanCollapseParentMargin(style, top: false) && children[children.Count - 1].HasCollapsibleMargins) {
            int lastIndex = children.Count - 1;
            HtmlRenderFlowBlock last = children[lastIndex];
            double childMargin = last.CollapsibleMarginBottom;
            style = style.Clone();
            collapsedBottomGroup = collapsedBottomGroup.Combine(last.CollapsibleMarginBottomGroup);
            style.MarginBottom = collapsedBottomGroup.Value;
            children[lastIndex] = last
                .AdjustTrailingFlowSpace(childMargin)
                .WithCollapsibleMargins(last.CollapsibleMarginTop, 0D, last.OwnerElement!,
                    topGroup: last.CollapsibleMarginTopGroup);
        }
        if (!usesVerticalBlockFormatting && children.Count > 0
            && (style.Display == "flow-root" || style.OverflowX is "auto" or "hidden" or "scroll"
                || style.OverflowY is "auto" or "hidden" or "scroll")) {
            int lastIndex = children.Count - 1;
            HtmlRenderFlowBlock last = children[lastIndex];
            if (last.OwnerElement != null
                && _inlineFloatOverhangs.TryGetValue(last.OwnerElement, out double floatOverhang)
                && _layoutStyles.TryGetValue(last.OwnerElement, out HtmlRenderBoxStyle? lastStyle)) {
                // The BFC contains the float, so its overhang can consume the final
                // block margin instead of extending the container below the float.
                children[lastIndex] = last.AdjustTrailingFlowSpace(Math.Min(Math.Max(0D, lastStyle.MarginBottom), floatOverhang));
            }
        }
        _layoutStyles[element] = style.Clone();

        if (usesVerticalBlockFormatting) {
            ArrangeVerticalBlockChildren(
                children,
                style,
                contentWidth,
                childPaintLayers,
                contentBreakOffsets,
                forcedBreaks,
                lineBreakGroups,
                continuationGroups,
                trailingGroups,
                runningStringAssignments,
                continuationBreakProgress,
                out contentHeight);
            AppendFlowPaintLayers(contentVisuals, hasLocallyPositionedPseudo
                ? childPaintLayers.Where(layer => !layer.Block.StackingZIndex.HasValue)
                : childPaintLayers);
        } else if (usesBlockFormatting) {
            string? childPageName = children.Count > 0 ? children[0].PageName : null;
            for (int childIndex = 0; childIndex < children.Count; childIndex++) {
                HtmlRenderFlowBlock child = children[childIndex];
                double childStart = contentHeight;
                if (childIndex > 0 && !string.Equals(childPageName, child.PageName, StringComparison.Ordinal)) {
                    forcedBreaks.Add(new HtmlRenderForcedBreak(childStart, HtmlPageBreakTarget.Page, child.PageName, changesPageName: true));
                }
                childPageName = child.PageName;
                if (child.BreakBefore != HtmlPageBreakTarget.None) {
                    forcedBreaks.Add(new HtmlRenderForcedBreak(childStart, child.BreakBefore));
                }
                foreach (HtmlRenderForcedBreak forcedBreak in child.ForcedBreaks) {
                    forcedBreaks.Add(forcedBreak.Translate(childStart));
                }
                if (child.HasCollapsibleMargins && child.CollapsibleMarginBottom > 0.0001D
                    && child.Height > child.CollapsibleMarginBottom + 0.0001D && child.OwnerElement != null) {
                    // A break after the child's last painted line may leave only
                    // its trailing margin to carry onto the next page.
                    continuationBreakProgress.Add(new HtmlInlineBreakProgress(
                        childStart + child.Height - child.CollapsibleMarginBottom,
                        0,
                        child.OwnerElement,
                        pageStartDiscardableMargin: child.CollapsibleMarginBottom,
                        isBlockExit: true));
                }
                if (childIndex > 0 && child.OwnerElement != null) {
                    continuationBreakProgress.Add(new HtmlInlineBreakProgress(
                        childStart,
                        0,
                        child.OwnerElement,
                        isBlockEntry: true,
                        pageStartDiscardableMargin: child.HasCollapsibleMargins
                            ? Math.Max(0D, child.CollapsibleMarginTop - child.LeadingFlowAdjustment)
                            : 0D));
                }
                childPaintLayers.Add(new FlowPaintLayer(child, 0D, childStart, childPaintLayers.Count));
                AppendKeepWithNextRange(children, childIndex, childStart, contentAvoidBreakRanges);

                contentHeight += child.Height;
                if (child.BreakAfter != HtmlPageBreakTarget.None) {
                    forcedBreaks.Add(new HtmlRenderForcedBreak(contentHeight, child.BreakAfter));
                }
                if (child.AvoidBreakInside)
                    contentAvoidBreakRanges.Add(new HtmlRenderAvoidBreakRange(childStart, contentHeight));
                contentAvoidBreakRanges.AddRange(child.AvoidBreakRanges.Select(range => range.Translate(childStart)));
                foreach (double offset in child.BreakOffsets) {
                    // A child's zero offset is its entry, not a break inside it. The
                    // preceding sibling already contributes that boundary; forwarding
                    // zero through a bordered parent can strand its top edge on the
                    // previous page before any child content fits.
                    // An anonymous float chunk can have a painted tail beyond its
                    // normal-flow height. Keep its internal opportunities, while
                    // the float context owns its final boundary and deferral.
                    bool isAnonymousFloatTail = child.OwnerElement == null
                        && child.PagedPaintExtent > child.Height + 0.0001D
                        && offset >= child.PagedPaintExtent - 0.0001D;
                    if (offset > 0.0001D && !isAnonymousFloatTail) contentBreakOffsets.Add(childStart + offset);
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
                foreach (HtmlCssRunningStringAssignment assignment in child.RunningStringAssignments) {
                    runningStringAssignments.Add(assignment.Translate(childStart));
                }
                foreach (HtmlInlineBreakProgress progress in child.InlineBreakProgress) {
                    continuationBreakProgress.Add(new HtmlInlineBreakProgress(
                        childStart + progress.Offset,
                        progress.LogicalCharacters,
                        progress.OwnerElement,
                        progress.IsBlockEntry,
                        progress.PageStartDiscardableMargin,
                        progress.IsBlockExit, progress.IsFlexGap));
                }

                contentBreakOffsets.Add(contentHeight);
            }
            AppendFlowPaintLayers(contentVisuals, hasLocallyPositionedPseudo
                ? childPaintLayers.Where(layer => !layer.Block.StackingZIndex.HasValue)
                : childPaintLayers);
            if (tag == "li" && children.Count > 0) {
                HtmlListMarker? marker = ResolveListMarker(element, style, contentWidth);
                if (marker?.IsOutside == true) {
                    AddOutsideMarkerForBlockChildren(contentVisuals, marker, contentWidth, style, element);
                }
            }
        } else {
            HtmlInlineLayout inline = inlineLayout!;
            if (IsVerticalWritingMode(style.WritingMode)) {
                inline = TransformSidewaysVerticalInlineLayout(inline, style, element);
            }
            inlineLayout = inline;
            if (emittedFloats != null && !EstablishesFloatContainingBlock(style)) {
                foreach (HtmlFloatExclusion exclusion in inline.FloatExclusions) emittedFloats.Add(exclusion);
            }
            if (inline.Height > inline.NormalFlowHeight + 0.0001D) {
                _inlineFloatOverhangs[element] = inline.Height - inline.NormalFlowHeight;
            }
            contentVisuals.AddRange(inline.Visuals);
            contentHeight = inline.Height;
            contentBreakOffsets.AddRange(inline.BreakOffsets);
            lineBreakOffsets.AddRange(inline.LineBreakOffsets);
            runningStringAssignments.AddRange(inline.RunningStringAssignments);
            forcedBreaks.AddRange(inline.ForcedBreaks);
            lineBreakGroups.AddRange(inline.LineBreakGroups);
            continuationGroups.AddRange(inline.ContinuationGroups);
            trailingGroups.AddRange(inline.TrailingGroups);
        }

        if (ownsFloatContext) contentHeight = Math.Max(contentHeight, contentFloats.Bottom);

        if (contentHeight <= 0D && style.ExplicitHeight == null && style.BackgroundColor == null && !style.HasBorderLayout) {
            contentHeight = tag == "div" || tag == "section" || tag == "article" ? 0D : style.LineHeight;
        }

        bool zeroHeightCollapsible = CanUseZeroHeightForMarginCollapse(style, parentStyle, contentHeight);
        double boxHeight = zeroHeightCollapsible ? 0D : ResolveBoxHeight(contentHeight, boxWidth, style);
        if (_inlineFloatOverhangs.TryGetValue(element, out double inlineOverhang)) {
            double normalFlowBoxHeight = ResolveBoxHeight(Math.Max(0D, contentHeight - inlineOverhang), boxWidth, style);
            double floatBottom = style.BorderTopWidth + style.PaddingTop + contentHeight;
            double boxOverhang = Math.Max(0D, floatBottom - normalFlowBoxHeight);
            bool containsOwnFloats = style.Display == "flow-root"
                || style.OverflowX is "auto" or "hidden" or "scroll"
                || style.OverflowY is "auto" or "hidden" or "scroll";
            if (boxOverhang > 0.0001D && !containsOwnFloats) _inlineFloatOverhangs[element] = boxOverhang;
            else _inlineFloatOverhangs.Remove(element);
        }
        double unclampedOuterHeight = style.MarginTop + boxHeight + style.MarginBottom;
        double outerHeight = unclampedOuterHeight;
        if (outerHeight <= 0D) outerHeight = 0.01D;
        var visuals = new List<HtmlRenderVisual>();
        var overflowContent = new List<HtmlRenderVisual>();
        var positionedRunningStringAssignments = new List<HtmlCssRunningStringAssignment>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        if (HtmlRenderSourceIdentity.TryGet(element, out string interactionSource)) {
            OfficeShape geometry = OfficeShape.Rectangle(Math.Max(0.01D, boxWidth), Math.Max(0.01D, boxHeight));
            geometry.FillColor = null;
            geometry.StrokeWidth = 0D;
            visuals.Add(new HtmlRenderShape(geometry, style.MarginLeft, style.MarginTop, visuals.Count, source: interactionSource));
        }
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop
            + ResolveButtonChildContentOffset(element, style, boxHeight, contentHeight);
        AppendBlockPositionedVisuals(
            element,
            Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
            Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
            style.MarginLeft + style.BorderLeftWidth,
            style.MarginTop + style.BorderTopWidth,
            PositionedPaintBand.Negative,
            style,
            hasLocallyPositionedPseudo ? childPaintLayers : null,
            contentX,
            contentY,
            overflowContent,
            positionedRunningStringAssignments);
        foreach (HtmlRenderVisual visual in contentVisuals) {
            overflowContent.Add(visual.Translate(contentX, contentY, overflowContent.Count));
        }
        if (style.Position != "static" || _localPositionedElements.ContainsKey(element)) {
            AppendBlockPositionedVisuals(
                element,
                Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
                Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
                style.MarginLeft + style.BorderLeftWidth,
                style.MarginTop + style.BorderTopWidth,
                PositionedPaintBand.NonNegative,
                style,
                hasLocallyPositionedPseudo ? childPaintLayers : null,
                contentX,
                contentY,
                overflowContent,
                positionedRunningStringAssignments);
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
        double contentYForBreaks = contentY;
        // Propagate only a child's explicit paged overhang. Ordinary child
        // geometry in a zero-height panel must not enlarge its print flow.
        double pagedPaintExtent = _options.Mode == HtmlRenderMode.Paged && style.OverflowY == "visible"
            && childPaintLayers.Any(layer => layer.Block.PagedPaintExtent > layer.Block.Height + 0.0001D)
            ? childPaintLayers.Aggregate(outerHeight, (extent, layer) =>
                Math.Max(extent, contentYForBreaks + layer.Y + layer.Block.PagedPaintExtent))
            : outerHeight;
        if (_options.Mode == HtmlRenderMode.Paged && style.OverflowY == "visible" && inlineLayout != null
            && inlineLayout.PagedPaintExtent > inlineLayout.Height + 0.0001D) {
            pagedPaintExtent = Math.Max(pagedPaintExtent, contentYForBreaks + inlineLayout.PagedPaintExtent);
        }
        HtmlRenderAvoidBreakRange? trailingBoxKeep = ResolveTrailingBoxKeepRange(style, contentHeight,
            outerHeight, contentYForBreaks, contentBreakOffsets, continuationBreakProgress);
        if (trailingBoxKeep.HasValue) contentAvoidBreakRanges.Add(trailingBoxKeep.Value);
        IEnumerable<double> breakOffsets = contentBreakOffsets.Select(offset => contentYForBreaks + offset)
            .Concat(new[] { outerHeight });
        bool hasPagedFloats = _options.Mode == HtmlRenderMode.Paged && contentFloats.HasFloats;
        if (pagedPaintExtent > outerHeight + 0.0001D || hasPagedFloats) {
            IReadOnlyList<(double Top, double Bottom)> atomicRanges = CollectAtomicParallelVisualRanges(visuals);
            breakOffsets = hasPagedFloats
                ? CollectSafeFloatBreaks(breakOffsets, contentFloats.GetFragmentPlacements(contentYForBreaks), atomicRanges,
                    pageBoundary?.PageHeight ?? _activePageGeometry.ContentHeight)
                : breakOffsets.Where(offset => !CrossesAtomicParallelVisual(atomicRanges, offset));
        }
        if (children.Count == 0 && contentVisuals.Count == 0) {
            breakOffsets = breakOffsets.Concat(CollectPositionedContainerBreakOffsets(
                element,
                Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
                Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
                style.MarginTop + style.BorderTopWidth,
                outerHeight));
        }
        IEnumerable<double> adjustedLineBreakOffsets = lineBreakOffsets.Select(offset => contentYForBreaks + offset);
        IEnumerable<HtmlRenderLineBreakGroup> adjustedLineBreakGroups = lineBreakGroups.Select(group => group.Translate(contentYForBreaks));
        IEnumerable<HtmlRenderContinuationGroup> adjustedContinuationGroups = continuationGroups.Select(group => group.Translate(contentX, contentYForBreaks));
        IEnumerable<HtmlRenderTrailingGroup> adjustedTrailingGroups = trailingGroups.Select(group =>
            group.Translate(
                contentX,
                contentYForBreaks,
                // A terminal footer inherits the box tail only when that cannot
                // move its source end before the body. A float may extend past
                // normal flow; clamping it there would rewind pagination.
                group.SourceEndsAt >= contentHeight - 0.0001D
                    && outerHeight >= contentYForBreaks + group.ContentEndsAt - 0.0001D
                    ? outerHeight : (double?)null));
        string? pageName = style.PageName;
        if (pageName == null && children.Count > 0) {
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
            pageName: pageName,
            unclampedHeight: unclampedOuterHeight,
            runningStringAssignments: runningStringAssignments
                .Select(assignment => assignment.Translate(contentYForBreaks))
                .Concat(positionedRunningStringAssignments)
                .OrderBy(assignment => assignment.OrderOffset),
            inlineBreakProgress: (inlineLayout?.BreakProgress ?? continuationBreakProgress).Select(progress =>
                new HtmlInlineBreakProgress(contentYForBreaks + progress.Offset, progress.LogicalCharacters, progress.OwnerElement, progress.IsBlockEntry, progress.PageStartDiscardableMargin, progress.IsBlockExit, progress.IsFlexGap)),
            inlineContinuationStart: ReferenceEquals(element, continuationTarget) ? continuationLogicalCharacters : 0,
            supportsInlineContinuationReflow: inlineLayout?.SupportsContinuationReflow == true
                || continuationBreakProgress.Any(progress => !progress.IsBlockExit && progress.OwnerElement != null),
            forcedBreaks: forcedBreaks.Select(item => item.Translate(contentYForBreaks)),
            avoidBreakRanges: contentAvoidBreakRanges.Select(range => range.Translate(contentYForBreaks)),
            pagedPaintExtent: pagedPaintExtent);
        block = ApplyElementSemantics(block, element, style);
        if (ownsFloatContext) foreach (HtmlRenderVisual visual in block.Visuals) visual.PaintPhase = HtmlRenderPaintPhase.Atomic;
        bool collapsesThrough = CanCollapseThroughEmptyBlock(style, usesBlockFormatting, children, contentVisuals, contentHeight);
        return StampViewport(AttachElementMargins(ApplyElementPositioning(block, style, containingWidth, containingHeight, element),
            style, element, collapsesThrough, collapsedTopGroup, collapsedBottomGroup));
    }

    private static HtmlRenderBoxStyle SuppressContinuationStartDecorations(HtmlRenderBoxStyle style) {
        HtmlRenderBoxStyle continuation = style.Clone();
        continuation.MarginTop = 0D;
        continuation.PaddingTop = 0D;
        continuation.Borders = new HtmlRenderBorderEdges(
            continuation.Borders.Top.WithWidth(0D),
            continuation.Borders.Right,
            continuation.Borders.Bottom,
            continuation.Borders.Left);
        continuation.BreakBefore = HtmlPageBreakTarget.None;
        continuation.StringSet = string.Empty;
        return continuation;
    }

    private HtmlRenderFlowBlock AttachElementMargins(HtmlRenderFlowBlock block, HtmlRenderBoxStyle style, IElement element,
        bool collapsesThrough = false, HtmlCollapsedMargin? topGroup = null, HtmlCollapsedMargin? bottomGroup = null) {
        HtmlRenderFlowBlock attached = block.WithCollapsibleMargins(style.MarginTop, style.MarginBottom, element,
            collapsesThrough, topGroup, bottomGroup);
        IReadOnlyList<HtmlCssRunningStringAssignment> ownAssignments =
            ResolveRunningStringAssignments(element, style, 0D);
        return ownAssignments.Count == 0
            ? attached
            : attached.WithRunningStringAssignments(ownAssignments.Concat(attached.RunningStringAssignments));
    }

    private IReadOnlyList<HtmlCssRunningStringAssignment> ResolveRunningStringAssignments(
        IElement element,
        HtmlRenderBoxStyle style,
        double offset) {
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        IReadOnlyList<HtmlCssRunningStringAssignment> ownAssignments = HtmlCssRunningStringParser.ResolveAssignments(
            element,
            style.StringSet,
            _options.MaxRunningStringCharacters,
            count => ChargeLayoutOperations(count, source),
            (contentElement, maximumCharacters, chargeOperations) =>
                ResolveRunningStringContentText(
                    contentElement,
                    style,
                    maximumCharacters,
                    chargeOperations),
            out bool limitExceeded);
        if (limitExceeded) {
            _diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.RunningStringLimitExceeded,
                "A CSS running-string value exceeded the configured character limit and was omitted.",
                HtmlDiagnosticSeverity.Warning,
                source,
                "limit=" + _options.MaxRunningStringCharacters);
        }
        int documentOrder = GetDocumentOrder(element);
        ownAssignments = ownAssignments
            .Select(assignment => assignment.InDocumentOrder(documentOrder))
            .ToList()
            .AsReadOnly();
        return Math.Abs(offset) <= 0.0001D
            ? ownAssignments
            : ownAssignments.Select(assignment => assignment.Translate(offset)).ToList().AsReadOnly();
    }

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
        if (style.ExplicitHeight.HasValue || style.AspectRatio.HasValue || style.MinHeight.HasValue && style.MinHeight.Value > 0D) return false;
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
        && !style.AspectRatio.HasValue
        && (!style.MinHeight.HasValue || style.MinHeight.Value <= 0D);

    private double FlushInlineNodes(ICollection<HtmlRenderFlowBlock> blocks, List<INode> nodes, double width, HtmlRenderBoxStyle style, IElement sourceElement, int depth, PagedFloatBoundary? pageBoundary = null, ICollection<HtmlFloatExclusion>? activeFloats = null, ICollection<HtmlFloatExclusion>? emittedFloats = null, double flowOffset = 0D, InlineFloatContext? floatContext = null) {
        if (nodes.Count == 0) return 0D;
        HtmlInlineLayout inline = LayoutInlineNodes(nodes, width, style, depth + 1, null, null,
            pageBoundary: pageBoundary,
            inheritedFloats: activeFloats?.Count > 0 ? activeFloats.Select(item => item.Shift(0D, -flowOffset)).ToArray() : null,
            applyTextIndent: blocks.Count == 0, floatContext: floatContext);
        if (activeFloats != null) {
            foreach (HtmlFloatExclusion exclusion in inline.FloatExclusions) {
                HtmlFloatExclusion positioned = exclusion.Shift(0D, flowOffset);
                activeFloats.Add(positioned);
                emittedFloats?.Add(positioned);
            }
        }
        if (inline.Height > inline.NormalFlowHeight + 0.0001D) {
            _inlineFloatOverhangs.TryGetValue(sourceElement, out double previousOverhang);
            _inlineFloatOverhangs[sourceElement] = Math.Max(
                previousOverhang,
                inline.Height - inline.NormalFlowHeight);
        }
        nodes.Clear();
        if (inline.InterruptedFlow != null) {
            blocks.Add(inline.InterruptedFlow);
            return inline.Height;
        }
        if (inline.Visuals.Count == 0) return 0D;
        // A floated sibling does not consume normal-flow height before the
        // following block. A block formatting context still contains its float.
        bool containsFloat = EstablishesFloatContainingBlock(style);
        double flowHeight = floatContext != null ? inline.NormalFlowHeight : activeFloats != null && (!containsFloat || HasMultiColumnLayout(style))
            ? inline.NormalFlowHeight : inline.Height;
        var block = new HtmlRenderFlowBlock(
            width,
            flowHeight,
            inline.Visuals,
            HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None,
            false,
            HtmlRenderStyleResolver.DescribeSource(sourceElement),
            inline.BreakOffsets,
            inline.LineBreakOffsets,
            style.Orphans,
            style.Widows,
            lineBreakGroups: inline.LineBreakGroups,
            continuationGroups: inline.ContinuationGroups,
            trailingGroups: inline.TrailingGroups,
            pageName: style.PageName,
            runningStringAssignments: inline.RunningStringAssignments,
            forcedBreaks: inline.ForcedBreaks,
            layoutViewportWidth: ActiveSurfaceWidth,
            layoutViewportHeight: _activePageGeometry.Height,
            pagedPaintExtent: inline.PagedPaintExtent);
        blocks.Add(block);
        return block.Height;
    }

    private static bool EstablishesFloatContainingBlock(HtmlRenderBoxStyle style) =>
        style.Display is "flow-root" or "flex" or "inline-flex" or "grid" or "inline-grid"
        || style.ColumnCount.HasValue || style.ColumnWidth.HasValue
        || style.OverflowX is "auto" or "hidden" or "scroll"
        || style.OverflowY is "auto" or "hidden" or "scroll";

    private bool HasBlockChildren(IElement element, double width, HtmlRenderBoxStyle parentStyle, int depth) {
        EnsureDepth(depth, element);
        if (HasBlockGeneratedContent(element, HtmlPseudoElementKind.Before, width, parentStyle)
            || HasBlockGeneratedContent(element, HtmlPseudoElementKind.After, width, parentStyle)) return true;
        foreach (IElement child in element.Children) {
            if (ShouldSkipElement(child)) continue;
            HtmlRenderBoxStyle style = _styleResolver.Resolve(child, width, parentStyle);
            if (style.FloatSide != "none") return true;
            if (style.Display == "contents" && HasBlockChildren(child, width, style, depth + 1)) return true;
            if (style.Display != "none" && ShouldExtractOutOfFlow(style) && !UsesInlineStaticPosition(child, style)) return true;
            if (style.Display != "none" && HtmlRenderStyleResolver.IsBlockElement(child, style)) return true;
            if (style.Display != "none" && ContainsFloatingDescendant(child, width, style, depth + 1)) return true;
        }

        return false;
    }

    private bool HasBlockGeneratedContent(IElement element, HtmlPseudoElementKind kind, double width, HtmlRenderBoxStyle parentStyle) {
        if (!_generatedContent.TryGetContent(element, kind, out _)
            || !_styleResolver.TryResolvePseudo(element, kind, width, parentStyle, out HtmlRenderBoxStyle style)) return false;
        if (CanPositionGeneratedContentLocally(style, parentStyle)) return false;
        return style.Display is "block" or "flow-root" or "list-item" or "flex" or "grid";
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
        if (style.MaxWidth.HasValue) width = Math.Min(width, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        if (style.MinWidth.HasValue) width = Math.Max(width, style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        return Math.Max(1D, width);
    }

    private HtmlRenderBoxStyle ResolveNormalFlowHorizontalAutoMargins(IElement element, HtmlRenderBoxStyle style, double containingWidth) {
        if (!style.MarginLeftAuto && !style.MarginRightAuto) return style;
        if (element.LocalName.Equals("table", StringComparison.OrdinalIgnoreCase) && !style.ExplicitWidth.HasValue) {
            // An auto-width table needs its intrinsic columns before its free space is known.
            return style;
        }

        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = IsReplacedImageElement(element)
            ? ResolveReplacedImageBoxWidth(element, style)
            : ResolveBoxWidth(availableWidth, style);
        double freeSpace = Math.Max(0D, containingWidth - style.MarginLeft - style.MarginRight - boxWidth);
        if (freeSpace <= 0D) return style;

        var resolved = style.Clone();
        if (style.MarginLeftAuto && style.MarginRightAuto) {
            resolved.MarginLeft += freeSpace / 2D;
            resolved.MarginRight += freeSpace / 2D;
        } else if (style.MarginLeftAuto) {
            resolved.MarginLeft += freeSpace;
        } else {
            resolved.MarginRight += freeSpace;
        }
        return resolved;
    }

    private static double ResolveBoxHeight(double contentHeight, double boxWidth, HtmlRenderBoxStyle style) {
        double height;
        bool usesAspectRatio = !style.ExplicitHeight.HasValue && style.AspectRatio.HasValue;
        if (style.ExplicitHeight.HasValue) {
            height = style.ExplicitHeight.Value;
        } else if (style.AspectRatio.HasValue) {
            double ratioWidth = style.BorderBox
                ? boxWidth
                : Math.Max(0D, boxWidth - style.HorizontalInsets);
            height = ratioWidth / style.AspectRatio.Value;
        } else {
            height = contentHeight;
        }
        if (!style.BorderBox || !style.ExplicitHeight.HasValue && !usesAspectRatio) height += style.VerticalInsets;
        if (style.MaxHeight.HasValue) height = Math.Min(height, style.MaxHeight.Value + (style.BorderBox ? 0D : style.VerticalInsets));
        if (style.MinHeight.HasValue) height = Math.Max(height, style.MinHeight.Value + (style.BorderBox ? 0D : style.VerticalInsets));
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

    private bool ShouldSkipElement(IElement element) {
        if (IsClosedDisclosureChild(element)) return true;
        string tag = element.TagName.ToLowerInvariant();
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

    private HtmlListMarker? ResolveListMarker(IElement element, HtmlRenderBoxStyle style, double containingWidth) {
        bool hasPseudoStyle = _styleResolver.TryResolvePseudo(
            element, HtmlPseudoElementKind.Marker, containingWidth, style, out HtmlRenderBoxStyle markerStyle);
        if (_generatedContent.Suppresses(element, HtmlPseudoElementKind.Marker)) return null;

        string? markerContent = null;
        if (_generatedContent.TryGetContent(element, HtmlPseudoElementKind.Marker, out HtmlGeneratedContent generatedContent)
            && generatedContent.Fragments.Any(fragment => fragment.Kind != HtmlGeneratedContentFragmentKind.Text)
            && TryCreateGeneratedListMarker(element, markerStyle, containingWidth, generatedContent, out HtmlRenderFlowBlock generatedMarkerBlock)) {
            return new HtmlListMarker(string.Empty, markerStyle, style.ListStylePosition, generatedMarkerBlock);
        }
        if (_generatedContent.TryGet(element, HtmlPseudoElementKind.Marker, out string generatedMarker)) {
            markerContent = ApplyTextTransform(generatedMarker, markerStyle);
        }
        if (markerContent == null && TryCreateListImageMarker(element, style, hasPseudoStyle ? markerStyle : style, containingWidth, out HtmlRenderFlowBlock imageMarker)) {
            return new HtmlListMarker(string.Empty, hasPseudoStyle ? markerStyle : style, style.ListStylePosition, imageMarker);
        }
        if (markerContent == null && string.Equals(style.ListStyleType, "none", StringComparison.OrdinalIgnoreCase)) return null;
        IElement? parent = element.ParentElement;
        if (markerContent == null && parent == null) markerContent = "• ";
        bool ordered = parent != null && string.Equals(parent.TagName, "ol", StringComparison.OrdinalIgnoreCase);
        string listStyle = style.ListStyleType.Length == 0 ? ordered ? "decimal" : "disc" : style.ListStyleType;
        int ordinal = HtmlListSemantics.TryResolveOrdinal(element, out int resolvedOrdinal) ? resolvedOrdinal : 1;
        if (markerContent == null && _counterStyles.TryFormatMarker(ordinal, listStyle, out string customMarker, out bool customMarkerLimited)) {
            if (customMarkerLimited) ReportCounterRepresentationLimit(element, listStyle);
            markerContent = customMarker.Length == 0 ? null : customMarker;
        }
        if (markerContent == null) {
            if (!HtmlCounterStyleFormatter.TryFormat(ordinal, listStyle, out string marker, out bool markerLimited)) {
                marker = ordered ? ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture) : "•";
            }
            if (markerLimited) ReportCounterRepresentationLimit(element, listStyle);
            if (marker.Length > 0) markerContent = marker + HtmlCounterStyleFormatter.MarkerSuffix(markerLimited ? "decimal" : listStyle);
        }
        if (string.IsNullOrEmpty(markerContent)) return null;
        return new HtmlListMarker(markerContent!, hasPseudoStyle ? markerStyle : style, style.ListStylePosition);
    }

    private bool TryCreateGeneratedListMarker(
        IElement element,
        HtmlRenderBoxStyle markerStyle,
        double containingWidth,
        HtmlGeneratedContent content,
        out HtmlRenderFlowBlock marker) {
        string source = DescribePseudoSource(element, HtmlPseudoElementKind.Marker);
        var runs = new List<HtmlInlineRun>();
        AddGeneratedInlineFragments(content, element, markerStyle, null, source, containingWidth, 0D, 0D, runs);
        runs = ApplyScopedFontFallbacks(runs);
        if (runs.Count == 0) {
            marker = null!;
            return false;
        }

        HtmlInlineLayout inline = LayoutInlineRuns(runs, Math.Max(1D, containingWidth * 0.5D), markerStyle, element);
        (double left, double top, double width, double height) = ResolveSemanticBounds(
            inline.Visuals,
            markerStyle.LineHeight,
            inline.Height);
        var label = new HtmlRenderSemanticGroup(
            HtmlRenderSemanticGroupRole.ListLabel,
            left,
            top,
            width,
            height,
            inline.Visuals,
            0,
            "list-marker");
        marker = new HtmlRenderFlowBlock(
            width,
            Math.Max(0.01D, inline.Height),
            new HtmlRenderVisual[] { label },
            HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None,
            true,
            source);
        return true;
    }

    private bool TryCreateListImageMarker(
        IElement element,
        HtmlRenderBoxStyle listStyle,
        HtmlRenderBoxStyle markerStyle,
        double containingWidth,
        out HtmlRenderFlowBlock marker) {
        marker = null!;
        if (string.Equals(listStyle.ListStyleImage, "none", StringComparison.OrdinalIgnoreCase)) return false;
        IReadOnlyList<string> sources = HtmlResourcePipeline.ExtractCssUrls(listStyle.ListStyleImage);
        if (sources.Count != 1) return false;
        string sourceDescription = HtmlRenderStyleResolver.DescribeSource(element) + "::marker:list-style-image";
        if (!TryResolveImageSource(sources[0], sourceDescription, out byte[]? bytes, out _, out OfficeImageInfo? imageInfo)
            || bytes == null || imageInfo == null) return false;
        IDocument? owner = element.Owner;
        if (owner == null) return false;
        IElement imageElement = owner.CreateElement("img");
        imageElement.SetAttribute("src", sources[0]);
        var imageStyle = new HtmlRenderBoxStyle {
            Display = "inline",
            PaintVisible = markerStyle.PaintVisible,
            Font = markerStyle.Font,
            Color = markerStyle.Color,
            LineHeight = markerStyle.LineHeight,
            SemanticRole = "list-marker",
            ApplyEmbeddedImageOrientation = listStyle.ApplyEmbeddedImageOrientation,
            ImageResolutionDpi = listStyle.ImageResolutionDpi
        };
        double imageWidth = ResolveFloatingImageOuterWidth(imageElement, imageStyle);
        HtmlRenderFlowBlock image = LayoutImage(imageElement, imageWidth, imageStyle);
        var markerVisual = new HtmlRenderSemanticGroup(
            HtmlRenderSemanticGroupRole.ListLabel,
            0D,
            0D,
            Math.Max(0.01D, image.Width),
            Math.Max(0.01D, image.Height),
            image.Visuals,
            0,
            "list-marker");
        marker = new HtmlRenderFlowBlock(
            image.Width,
            image.Height,
            new[] { markerVisual },
            HtmlPageBreakTarget.None,
            HtmlPageBreakTarget.None,
            true,
            sourceDescription);
        return true;
    }

    private void ReportCounterRepresentationLimit(IElement element, string style) {
        AddUnsupported(
            HtmlRenderDiagnosticCodes.CounterRepresentationLimitExceeded,
            "A CSS counter representation exceeded the managed rendering budget and used a decimal fallback.",
            element,
            "list-style-type=" + style,
            OfficeConversionLossKind.Approximation);
    }
}
