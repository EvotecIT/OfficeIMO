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
        if (_deferredColumnWrappers.Contains(element)) {
            // This resolved style belongs to the current element's layout pass.
            // Keep its flex/grid placement and its painted entry in agreement.
            style.MarginTop = 0D;
            style.BreakBefore = HtmlPageBreakTarget.Page;
        }
        bool continuesThisBox = continuationTarget != null
            && ContainsElementOrSelf(element, continuationTarget)
            && (!ReferenceEquals(element, continuationTarget) || continuationLogicalCharacters > 0);
        // Page and continuation reflow can enter here with a fresh style,
        // bypassing normal child placement and its auto-margin preparation.
        if (style.HasIntrinsicWidths && !style.IntrinsicWidthsResolved) {
            style = ResolveOrdinaryIntrinsicWidths(element, style, containingWidth, depth);
            style = ResolveNormalFlowHorizontalAutoMargins(element, style, containingWidth);
        }
        style = ResolveColumnFragmentStyle(element, PrepareButtonChildStyle(element, style));
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
        RememberColumnSourceStyle(element, style);
        using OrdinaryPageSourceScope pageSourceScope = EnterOrdinaryPageSource(style);
        string tag = element.TagName.ToLowerInvariant();
        double? containingHeight = ResolveContainingBlockHeight(parentStyle);
        if (IsReplacedImageElementTag(tag)) return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(ApplySpecializedElementSemantics(LayoutImage(element, containingWidth, style), element, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        if (tag == "math" && TryLayoutMath(element, containingWidth, style, inheritedLink: null, shrinkToFit: false, out HtmlRenderFlowBlock mathBlock, out _)) return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(ApplySpecializedElementSemantics(mathBlock, element, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        if (IsFormControlElement(tag) && !UsesButtonChildLayout(element)) return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(ApplySpecializedElementSemantics(LayoutFormControl(element, containingWidth, style), element, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        if (style.Display is "table" or "inline-table") {
            HtmlRenderFlowBlock table = LayoutTable(element, containingWidth, style, depth, continuationTarget);
            // The table envelope owns its used border box before optional semantic
            // wrapping. Its fixed columns can enlarge the authored width.
            double tableWidth = table.Visuals.OfType<HtmlRenderSemanticGroup>().First().Width;
            return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(
                ApplySpecializedElementSemantics(table, element, style), style, element, containingWidth, tableWidth),
                style, containingWidth, containingHeight, element), style, element));
        }
        if (tag == "hr") return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(ApplySpecializedElementSemantics(LayoutHorizontalRule(element, containingWidth, style), element, style), style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        if (style.Display == "flex" && TryLayoutFlexContainer(element, containingWidth, style, depth, continuationTarget, pageBoundary, out HtmlRenderFlowBlock flexBlock)) {
            flexBlock = ApplyElementSemantics(flexBlock, element, style);
            return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(flexBlock, style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        }
        if (style.Display == "grid" && TryLayoutGridContainer(element, containingWidth, style, depth, pageBoundary, out HtmlRenderFlowBlock gridBlock)) {
            gridBlock = ApplyElementSemantics(gridBlock, element, style);
            return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(gridBlock, style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        }
        if (TryLayoutMultiColumnContainer(element, containingWidth, style, depth, out HtmlRenderFlowBlock columnsBlock)) {
            columnsBlock = ApplyElementSemantics(columnsBlock, element, style);
            return StampViewport(AttachElementMargins(ApplyElementPositioning(ApplyOverflowToSpecializedBlock(columnsBlock, style, element, containingWidth), style, containingWidth, containingHeight, element), style, element));
        }

        using ConstrainedProjectionScope projectionScope = EnterConstrainedProjection(element, style);
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
        var contentSafeBreakRanges = new List<HtmlRenderBreakRange>();
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
            HtmlListMarker? marker = tag == "li" && IsFirstColumnSourceFragment(element) ? ResolveListMarker(element, style, contentWidth) : null;
            int skipped = ReferenceEquals(element, continuationTarget) ? continuationLogicalCharacters : 0;
            double extent = IsVerticalWritingMode(style.WritingMode)
                ? ResolveVerticalInlineExtent(style, parentStyle, contentWidth) : contentWidth;
            inlineLayout = LayoutInlineNodes(SelectedColumnChildNodes(element), extent, style, depth, marker, element, skipped, pageBoundary?.Shift(style.MarginTop + style.BorderTopWidth + style.PaddingTop), inheritedFloats, floatContext: contentFloats);
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
        RememberColumnSourceStyle(element, style);

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
            object[] childBoundaries = CreateFlowBoundaries(children);
            for (int childIndex = 0; childIndex < children.Count; childIndex++) {
                HtmlRenderFlowBlock child = children[childIndex];
                double childStart = contentHeight;
                if (childIndex > 0 && !string.Equals(childPageName, child.PageName, StringComparison.Ordinal)) {
                    forcedBreaks.Add(new HtmlRenderForcedBreak(childStart, HtmlPageBreakTarget.Page, child.PageName, changesPageName: true));
                }
                childPageName = child.PageName;
                AppendChildForcedBreaks(children, childIndex, childStart, forcedBreaks, childBoundaries);
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

                contentHeight += child.NormalFlowAdvance;
                if (child.AvoidBreakInsideScope != HtmlFragmentationScope.None)
                    contentAvoidBreakRanges.Add(new HtmlRenderAvoidBreakRange(childStart, contentHeight, Scope: child.AvoidBreakInsideScope));
                contentAvoidBreakRanges.AddRange(child.AvoidBreakRanges.Select(range => range.Translate(childStart)));
                contentSafeBreakRanges.AddRange(child.SafeBreakRanges.Select(range => range.Translate(childStart)));
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
            if (tag == "li" && children.Count > 0 && IsFirstColumnSourceFragment(element)) {
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
            contentSafeBreakRanges.AddRange(inline.SafeBreakRanges);
            lineBreakOffsets.AddRange(inline.LineBreakOffsets);
            runningStringAssignments.AddRange(inline.RunningStringAssignments);
            forcedBreaks.AddRange(inline.ForcedBreaks);
            lineBreakGroups.AddRange(inline.LineBreakGroups);
            contentAvoidBreakRanges.AddRange(inline.AvoidBreakRanges);
            continuationGroups.AddRange(inline.ContinuationGroups);
            trailingGroups.AddRange(inline.TrailingGroups);
        }

        if (ownsFloatContext) contentHeight = Math.Max(contentHeight, contentFloats.Bottom);

        HtmlRenderFlowBlock? projectedContent = null;
        double pagedEntryGap = 0D;
        if (projectionScope.CanProject && pageBoundary.HasValue) {
            double sourceExtent = Math.Max(contentHeight, inlineLayout?.PagedPaintExtent ?? 0D);
            foreach (FlowPaintLayer child in childPaintLayers) {
                CheckCancellation();
                ChargeLayoutOperation("ordinary source projection extent");
                sourceExtent = Math.Max(sourceExtent, child.Y + ResolveNormalFlowPaintExtent(child.Block));
            }
            ChargeLayoutOperations(1L + contentVisuals.Count + contentBreakOffsets.Count + lineBreakOffsets.Count
                + lineBreakGroups.Count + continuationGroups.Count + trailingGroups.Count
                + runningStringAssignments.Count + continuationBreakProgress.Count + forcedBreaks.Count
                + contentAvoidBreakRanges.Count + contentSafeBreakRanges.Count,
                "ordinary source projection metadata");
            var source = new HtmlRenderFlowBlock(contentWidth, Math.Max(0D, contentHeight), contentVisuals,
                HtmlPageBreakTarget.None, HtmlPageBreakTarget.None, style.AvoidBreakInside,
                HtmlRenderStyleResolver.DescribeSource(element), contentBreakOffsets, lineBreakOffsets,
                style.Orphans, style.Widows, lineBreakGroups, continuationGroups, trailingGroups,
                unclampedHeight: contentHeight, runningStringAssignments: runningStringAssignments,
                inlineBreakProgress: inlineLayout?.BreakProgress ?? continuationBreakProgress,
                forcedBreaks: forcedBreaks, avoidBreakRanges: contentAvoidBreakRanges,
                pagedPaintExtent: sourceExtent, safeBreakRanges: contentSafeBreakRanges);
            projectedContent = ProjectOrdinaryContent(source,
                pageBoundary.Value.Shift(style.MarginTop + style.BorderTopWidth + style.PaddingTop), out pagedEntryGap);
            contentVisuals.Clear(); contentVisuals.AddRange(projectedContent.Visuals);
            contentHeight = projectedContent.NormalFlowAdvance;
            contentBreakOffsets.Clear(); contentBreakOffsets.AddRange(projectedContent.BreakOffsets);
            contentSafeBreakRanges.Clear(); contentSafeBreakRanges.AddRange(projectedContent.SafeBreakRanges);
            contentAvoidBreakRanges.Clear(); contentAvoidBreakRanges.AddRange(projectedContent.AvoidBreakRanges);
            lineBreakOffsets.Clear();
            lineBreakGroups.Clear(); lineBreakGroups.AddRange(projectedContent.LineBreakGroups);
            continuationGroups.Clear(); continuationGroups.AddRange(projectedContent.ContinuationGroups);
            trailingGroups.Clear(); trailingGroups.AddRange(projectedContent.TrailingGroups);
            runningStringAssignments.Clear(); runningStringAssignments.AddRange(projectedContent.RunningStringAssignments);
            continuationBreakProgress.Clear(); continuationBreakProgress.AddRange(projectedContent.InlineBreakProgress);
            forcedBreaks.Clear(); forcedBreaks.AddRange(projectedContent.ForcedBreaks);
        }

        if (contentHeight <= 0D && style.ExplicitHeight == null && style.BackgroundColor == null && !style.HasBorderLayout) {
            contentHeight = tag == "div" || tag == "section" || tag == "article" ? 0D : style.LineHeight;
        }

        bool hasZeroHeightOrdinaryContent = contentHeight == 0D
            && (inlineLayout != null && inlineLayout.Height == 0D && inlineLayout.HasInFlowLineBoxes == true
                || children.Count > 0 && children.All(child => child.Height == 0D)
                    && children.Any(child => child.HasZeroHeightOrdinaryContent));
        bool zeroHeightCollapsible = CanUseZeroHeightForMarginCollapse(style, parentStyle, contentHeight);
        double boxHeight = zeroHeightCollapsible ? 0D : ResolveBoxHeight(contentHeight, boxWidth, style,
            allowZeroHeight: contentHeight <= 0D && (style.Display == "block"
                || style.Display == "flow-root" && hasZeroHeightOrdinaryContent));
        bool hasOrdinaryFragmentSource = contentVisuals.Any(HasOrdinaryFragmentSource);
        bool retainsConstrainedPaint = pageSourceScope.RetainsConstrainedPaint && _columnContentDepth == 0
            && !IsAtomicPaginationOwner(style) && HasEffectiveConstrainedBlockHeight(style, boxHeight, contentHeight);
        bool ownsConstrainedSource = pageSourceScope.IsOrdinary && retainsConstrainedPaint;
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
        // Margins change normal progression, not the nonnegative allocation
        // used by the fragment planner. Keep existing paint/float carriers while
        // retaining the signed normal advance independently.
        bool hasContentCarrier = contentVisuals.Any(ContainsInFlowPaint) || inlineLayout?.HasInFlowLineBoxes == true
            || inlineLayout?.FloatExclusions.Count > 0 || hasLocallyPositionedPseudo;
        double outerHeight = Math.Max(0D, unclampedOuterHeight);
        // A float-only zero-height box retains its existing pagination carrier.
        // A real box cancelled by negative margins has no such allocation.
        bool retainsZeroHeightFlowCarrier = unclampedOuterHeight == 0D && boxHeight == 0D
            && hasContentCarrier && !hasZeroHeightOrdinaryContent;
        if (retainsZeroHeightFlowCarrier) outerHeight = HtmlRenderFlowBlock.ZeroHeightFlowCarrierAdvance;
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
        // Auto-height wrappers import explicit overhang. An authored height
        // constraint also retains its in-flow content beyond the used box.
        double pagedPaintExtent = _options.Mode == HtmlRenderMode.Paged && style.OverflowY == "visible"
            && childPaintLayers.Any(layer => layer.Block.PagedPaintExtent > layer.Block.Height + 0.0001D)
            ? childPaintLayers.Aggregate(outerHeight, (extent, layer) =>
                Math.Max(extent, contentYForBreaks + layer.Y + layer.Block.PagedPaintExtent))
            : outerHeight;
        if (style.MarginBottom < 0D)
            pagedPaintExtent = Math.Max(pagedPaintExtent, style.MarginTop + boxHeight);
        if (_options.Mode == HtmlRenderMode.Paged && style.OverflowY == "visible" && inlineLayout != null
            && inlineLayout.PagedPaintExtent > inlineLayout.Height + 0.0001D) {
            pagedPaintExtent = Math.Max(pagedPaintExtent, contentYForBreaks + inlineLayout.PagedPaintExtent);
        }
        if (!usesVerticalBlockFormatting) {
            pagedPaintExtent = ResolveConstrainedBlockPaintExtent(style, retainsConstrainedPaint, pagedPaintExtent,
                contentYForBreaks, contentHeight, inlineLayout, childPaintLayers);
            AppendConstrainedBlockAllocationTail(style, ownsConstrainedSource && hasOrdinaryFragmentSource
                && !contentFloats.HasFloats && !visuals.Any(ContainsPositionedFlowPaint), boxHeight, contentYForBreaks,
                contentHeight, childPaintLayers, contentSafeBreakRanges,
                pageBoundary?.PageHeight ?? _activePageGeometry.ContentHeight);
        }
        double columnOverflowExtent = outerHeight;
        if (_columnContentDepth > 0 && style.OverflowY == "visible" && !usesVerticalBlockFormatting) {
            columnOverflowExtent = Math.Max(columnOverflowExtent, ResolveColumnStaticPositionPaintExtent(visuals));
            if (style.ExplicitHeight.HasValue || style.MaxHeight.HasValue || style.AspectRatio.HasValue)
                columnOverflowExtent = Math.Max(columnOverflowExtent, contentYForBreaks + contentHeight);
            columnOverflowExtent = AppendColumnChildOverflowFacts(childPaintLayers, contentYForBreaks,
                columnOverflowExtent, contentBreakOffsets, contentSafeBreakRanges,
                contentAvoidBreakRanges, forcedBreaks);
        }
        double? sourceContentHeight = !usesVerticalBlockFormatting && children.Any(child => child.UnfragmentedHeight.HasValue)
            ? children.Sum(child => child.UnfragmentedHeight ?? child.Height) : null;
        if (sourceContentHeight.HasValue && ownsFloatContext)
            sourceContentHeight = Math.Max(sourceContentHeight.Value, contentFloats.Bottom);
        // A projected column set can leave a retained source-allocation tail.
        // Its last painted line need not travel with the wrapper's bottom edge.
        HtmlRenderAvoidBreakRange? trailingBoxKeep = ResolveTrailingBoxKeepRange(style,
            Math.Max(contentHeight, sourceContentHeight ?? 0D),
            outerHeight, contentYForBreaks, contentBreakOffsets, continuationBreakProgress);
        if (trailingBoxKeep.HasValue) contentAvoidBreakRanges.Add(trailingBoxKeep.Value);
        IEnumerable<double> breakOffsets = contentBreakOffsets.Select(offset => contentYForBreaks + offset)
            .Concat(new[] { outerHeight });
        IEnumerable<HtmlRenderBreakRange> safeBreakRanges = contentSafeBreakRanges.Select(range => range.Translate(contentYForBreaks));
        bool hasPagedFloats = _options.Mode == HtmlRenderMode.Paged && contentFloats.HasFloats;
        if (pagedPaintExtent > outerHeight + 0.0001D || hasPagedFloats) {
            IReadOnlyList<(double Top, double Bottom)> atomicRanges = CollectAtomicParallelVisualRanges(visuals);
            IReadOnlyList<InlineFloatPlacement> fragmentPlacements = hasPagedFloats
                ? contentFloats.GetFragmentPlacements(contentYForBreaks) : Array.Empty<InlineFloatPlacement>();
            breakOffsets = hasPagedFloats
                ? CollectSafeFloatBreaks(breakOffsets, fragmentPlacements, atomicRanges,
                    pageBoundary?.PageHeight ?? _activePageGeometry.ContentHeight)
                : breakOffsets.Where(offset => !CrossesAtomicParallelVisual(atomicRanges, offset));
            safeBreakRanges = hasPagedFloats
                ? CollectSafeFloatBreakRanges(safeBreakRanges, fragmentPlacements, atomicRanges,
                    pageBoundary?.PageHeight ?? _activePageGeometry.ContentHeight)
                : HtmlRenderBreakRange.Exclude(safeBreakRanges,
                    atomicRanges.Select(range => new HtmlRenderBreakRange(range.Top, range.Bottom)));
        }
        if (contentSafeBreakRanges.Count > 0) {
            // A short fixed-height child can paint over later normal-flow
            // siblings without extending this wrapper's final flow height.
            // Its blank source interval cannot cut those sibling carriers.
            safeBreakRanges = ConstrainParallelPaintBreakRanges(
                HtmlRenderBreakRange.Normalize(safeBreakRanges, Math.Max(pagedPaintExtent, columnOverflowExtent)), visuals);
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

        if (columnOverflowExtent > outerHeight + 0.0001D)
            RegisterColumnOverflowOpacity(element, style, depth);
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
            inlineBreakProgress: (projectedContent?.InlineBreakProgress ?? inlineLayout?.BreakProgress ?? continuationBreakProgress).Select(progress =>
                new HtmlInlineBreakProgress(contentYForBreaks + progress.Offset, progress.LogicalCharacters, progress.OwnerElement, progress.IsBlockEntry, progress.PageStartDiscardableMargin, progress.IsBlockExit, progress.IsFlexGap)),
            inlineContinuationStart: ReferenceEquals(element, continuationTarget) ? continuationLogicalCharacters : 0,
            supportsInlineContinuationReflow: inlineLayout?.SupportsContinuationReflow == true
                || continuationBreakProgress.Any(progress => !progress.IsBlockExit && progress.OwnerElement != null),
            forcedBreaks: forcedBreaks.Select(item => item.Translate(contentYForBreaks)),
            avoidBreakRanges: contentAvoidBreakRanges.Select(range => range.Translate(contentYForBreaks)),
            pagedPaintExtent: Math.Max(pagedPaintExtent, columnOverflowExtent),
            safeBreakRanges: safeBreakRanges);
        block.RetainsZeroHeightFlowCarrier = retainsZeroHeightFlowCarrier;
        if (sourceContentHeight.HasValue) {
            block.UnfragmentedHeight = Math.Max(0D, style.MarginTop + style.MarginBottom
                + ResolveBoxHeight(sourceContentHeight.Value, boxWidth, style));
        }
        if (ownsFloatContext) block = OrderBoxFlowPaint(block);
        block = ApplyElementSemantics(block, element, style);
        if (ownsFloatContext) CloseBoxFlowPaint(block);
        bool collapsesThrough = CanCollapseThroughEmptyBlock(style, usesBlockFormatting, children,
            contentVisuals, contentHeight, inlineLayout?.HasInFlowLineBoxes == true);
        HtmlRenderFlowBlock completed = StampViewport(AttachElementMargins(ApplyElementPositioning(block, style, containingWidth, containingHeight, element),
            style, element, collapsesThrough, collapsedTopGroup, collapsedBottomGroup));
        if (pagedEntryGap > 0.0001D) completed = completed.AdjustLeadingFlowSpace(-pagedEntryGap);
        if (projectionScope.CanProjectAtomic && pageBoundary.HasValue)
            completed = ProjectAtomicOrdinaryBlock(completed, pageBoundary.Value,
                style.MarginTop, style.MarginTop + boxHeight);
        RememberConstrainedOverflow(element, ownsConstrainedSource && hasOrdinaryFragmentSource, completed);
        if (columnOverflowExtent > completed.Height + 0.0001D) {
            HtmlRenderFlowBlock overflow = completed.ForPagination();
            completed = completed.WithVisuals(completed.Visuals, pagedPaintExtent: pagedPaintExtent);
            completed.ColumnOverflowContent = overflow;
        }
        completed.KnownNoInFlowLineBoxes = !usesBlockFormatting && inlineLayout?.HasInFlowLineBoxes == false;
        completed.HasZeroHeightOrdinaryContent = hasZeroHeightOrdinaryContent;
        RecordColumnSourceFragmentSize(element, completed, style);
        return completed;
    }

}
