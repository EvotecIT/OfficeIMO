using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<HtmlRenderFlowBlock> BuildChildBlocks(
        IElement container,
        double width,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        IElement? continuationTarget = null,
        int continuationLogicalCharacters = 0,
        PagedFloatBoundary? pageBoundary = null,
        IReadOnlyList<HtmlFloatExclusion>? inheritedFloats = null,
        ICollection<HtmlFloatExclusion>? emittedFloats = null, InlineFloatContext? floatContext = null) =>
        BuildChildBlocks(
            container,
            container.ChildNodes,
            width,
            parentStyle,
            depth,
            includeGeneratedBefore: true,
            includeGeneratedAfter: true,
            continuationTarget,
            continuationLogicalCharacters,
            pageBoundary,
            inheritedFloats,
            emittedFloats, floatContext);

    private IReadOnlyList<HtmlRenderFlowBlock> BuildChildBlocks(
        IElement container,
        IEnumerable<INode> nodes,
        double width,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        bool includeGeneratedBefore,
        bool includeGeneratedAfter,
        IElement? continuationTarget = null,
        int continuationLogicalCharacters = 0,
        PagedFloatBoundary? pageBoundary = null,
        IReadOnlyList<HtmlFloatExclusion>? inheritedFloats = null,
        ICollection<HtmlFloatExclusion>? emittedFloats = null, InlineFloatContext? floatContext = null) {
        EnsureDepth(depth, container);
        bool ownsFloatContext = floatContext == null;
        floatContext ??= new InlineFloatContext(width, inheritedFloats);
        if (container.TagName.Equals("details", StringComparison.OrdinalIgnoreCase) && !container.HasAttribute("open")) {
            IElement? summary = container.Children.FirstOrDefault(child =>
                child.TagName.Equals("summary", StringComparison.OrdinalIgnoreCase));
            nodes = nodes.Where(node => ReferenceEquals(node, summary));
        }
        var blocks = new List<HtmlRenderFlowBlock>();
        IElement? continuationChild = FindDirectChildContaining(container, continuationTarget);
        bool seekingContinuation = continuationChild != null;
        if (includeGeneratedBefore && !seekingContinuation) AddGeneratedContentBlock(blocks, container, HtmlPseudoElementKind.Before, width, parentStyle);
        double flowHeight = blocks.Sum(block => block.Height);
        var adjoiningMargins = new AdjoiningMarginState();
        var inlineNodes = new List<INode>();
        List<HtmlFloatExclusion>? activeFloats = pageBoundary.HasValue || emittedFloats != null
            ? inheritedFloats == null ? new List<HtmlFloatExclusion>() : new List<HtmlFloatExclusion>(inheritedFloats)
            : null;
        foreach (TableFlowEntry entry in EnumerateTableFlowEntries(container, nodes, width, parentStyle)) {
            if (entry.TableNodes != null) {
                if (seekingContinuation && !entry.TableNodes.OfType<IElement>().Any(element => ReferenceEquals(element, continuationChild))) continue;
                seekingContinuation = false;
                double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth,
                    pageBoundary?.Shift(flowHeight), activeFloats, emittedFloats, flowHeight, floatContext.At(width, 0D, flowHeight));
                flowHeight += inlineHeight;
                HtmlRenderBoxStyle tableStyle = CreateAnonymousBoxStyle(parentStyle, "table", "anonymous-table");
                double tableY = flowHeight;
                tableStyle = PlaceIndependentBlockBesideFloats(tableStyle, width, floatContext, ref tableY);
                if (tableY > flowHeight) {
                    blocks.Add(CreateFloatClearanceBlock(width, tableY - flowHeight));
                    flowHeight = tableY;
                }
                TableFormattingStructure formatting = BuildTableFormattingStructure(container, width, tableStyle, depth, entry.TableNodes);
                HtmlRenderFlowBlock tableBlock = LayoutTable(container, width, tableStyle, depth, continuationTarget, formatting);
                blocks.Add(tableBlock);
                flowHeight += tableBlock.Height;
                if (continuationTarget != null && formatting.Rows.Any(row => row.Contains(continuationTarget))) {
                    continuationTarget = null;
                    continuationLogicalCharacters = 0;
                }
                adjoiningMargins.Clear();
                continue;
            }
            INode node = entry.Node!;
            CheckCancellation();
            for (int index = (activeFloats?.Count ?? 0) - 1; index >= 0; index--) {
                if (activeFloats![index].Bottom <= flowHeight + 0.0001D) activeFloats.RemoveAt(index);
            }
            if (IsClosedDisclosureChild(node)) continue;
            if (seekingContinuation) {
                if (node is not IElement candidate || !ReferenceEquals(candidate, continuationChild)) continue;
                seekingContinuation = false;
            }
            if (node is IElement element) {
                if (ShouldSkipElement(element)) {
                    continue;
                }

                HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(element, width, parentStyle);
                childStyle = ForwardTableCellHeightBasis(element, childStyle, parentStyle);
                if (childStyle.Display == "none") {
                    continue;
                }
                if (HtmlCssRunningElementParser.TryParsePosition(childStyle.Position, out string runningElementName)) {
                    double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth, pageBoundary?.Shift(flowHeight), activeFloats, emittedFloats, flowHeight, floatContext.At(width, 0D, flowHeight));
                    flowHeight += inlineHeight;
                    if (inlineHeight > 0D) adjoiningMargins.Clear();
                    IReadOnlyList<HtmlCssRunningStringAssignment> assignments = CaptureRunningElement(
                        element,
                        runningElementName,
                        width,
                        childStyle,
                        parentStyle,
                        depth + 1);
                    blocks.Add(new HtmlRenderFlowBlock(
                        width,
                        0D,
                        Array.Empty<HtmlRenderVisual>(),
                        HtmlPageBreakTarget.None,
                        HtmlPageBreakTarget.None,
                        false,
                        HtmlRenderStyleResolver.DescribeSource(element),
                        runningStringAssignments: assignments,
                        layoutViewportWidth: ActiveSurfaceWidth,
                        layoutViewportHeight: _activePageGeometry.Height));
                    adjoiningMargins.Clear();
                    continue;
                }
                FlattenedSemanticBoundary? flattenedSemanticBoundary = childStyle.Display == "contents"
                    ? CreateFlattenedSemanticBoundary(element, childStyle)
                    : null;
                if (childStyle.Display == "contents" && HasBlockChildren(element, width, childStyle, depth + 1)) {
                    double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth, pageBoundary?.Shift(flowHeight), activeFloats, emittedFloats, flowHeight, floatContext.At(width, 0D, flowHeight));
                    flowHeight += inlineHeight;
                    bool carriesContinuation = ContainsElementOrSelf(element, continuationTarget);
                    List<HtmlFloatExclusion>? flattenedFloats = activeFloats != null
                        ? new List<HtmlFloatExclusion>()
                        : null;
                    IReadOnlyList<HtmlRenderFlowBlock> flattenedBlocks = BuildChildBlocks(
                        element,
                        width,
                        childStyle,
                        depth + 1,
                        carriesContinuation ? continuationTarget : null,
                        carriesContinuation ? continuationLogicalCharacters : 0,
                        pageBoundary?.Shift(flowHeight),
                        activeFloats?.Count > 0 ? activeFloats.Select(item => item.Shift(0D, -flowHeight)).ToArray() : null,
                        flattenedFloats, floatContext.At(width, 0D, flowHeight));
                    if (flattenedFloats != null) {
                        foreach (HtmlFloatExclusion exclusion in flattenedFloats) {
                            HtmlFloatExclusion positioned = exclusion.Shift(0D, flowHeight);
                            activeFloats?.Add(positioned);
                            emittedFloats?.Add(positioned);
                        }
                    }
                    foreach (HtmlRenderFlowBlock flattenedBlock in ApplyFlattenedElementSemantics(flattenedBlocks, flattenedSemanticBoundary!)) {
                        blocks.Add(flattenedBlock);
                        flowHeight += flattenedBlock.Height;
                    }
                    if (carriesContinuation) {
                        continuationTarget = null;
                        continuationLogicalCharacters = 0;
                    }

                    adjoiningMargins.Clear();
                    continue;
                }
                if (ShouldExtractOutOfFlow(childStyle)) {
                    if (UsesInlineStaticPosition(element, childStyle)) {
                        inlineNodes.Add(node);
                        continue;
                    }
                    double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth, pageBoundary?.Shift(flowHeight), activeFloats, emittedFloats, flowHeight, floatContext.At(width, 0D, flowHeight));
                    flowHeight += inlineHeight;
                    if (inlineHeight > 0D) {
                        adjoiningMargins.Clear();
                    }
                    PositionedStaticAnchor? staticAnchor = HtmlRenderStyleResolver.IsBlockElement(element, childStyle)
                        ? new PositionedStaticAnchor(container, 0D, flowHeight)
                        : null;
                    RegisterOutOfFlowElement(container, element, childStyle, parentStyle, depth + 1, staticAnchor);
                    continue;
                }

                if (childStyle.FloatSide != "none" || TryGetPageFloatSide(element, childStyle, out _)
                    || TryGetColumnEdgeFloatSide(element, childStyle, out _, out _)) {
                    inlineNodes.Add(node);
                    continue;
                }

                if (HtmlRenderStyleResolver.IsBlockElement(element, childStyle)) {
                    double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth, pageBoundary?.Shift(flowHeight), activeFloats, emittedFloats, flowHeight, floatContext.At(width, 0D, flowHeight));
                    flowHeight += inlineHeight;
                    if (inlineHeight > 0D) {
                        adjoiningMargins.Clear();
                    }
                    if (activeFloats == null) {
                        double blockY = Math.Max(flowHeight, floatContext.Clearance(childStyle.ClearSide));
                        childStyle = PlaceIndependentBlockBesideFloats(childStyle, width, floatContext, ref blockY);
                        if (blockY > flowHeight) {
                            blocks.Add(CreateFloatClearanceBlock(width, blockY - flowHeight));
                            flowHeight = blockY;
                            adjoiningMargins.Clear();
                        }
                    }
                    var anticipatedMargins = adjoiningMargins;
                    anticipatedMargins.Add(new HtmlCollapsedMargin(childStyle.MarginTop));
                    double anticipatedMarginAdjustment = adjoiningMargins.Count > 0
                        ? anticipatedMargins.Allocated - anticipatedMargins.Collapsed : 0D;
                    HtmlRenderBoxStyle unavoidedStyle = childStyle.Clone();
                    childStyle = AvoidActiveFloatsForFormattingContext(
                        unavoidedStyle, width, flowHeight, activeFloats);
                    childStyle = ResolveOrdinaryIntrinsicWidths(element, childStyle, width, depth + 1);
                    childStyle = ResolveNormalFlowHorizontalAutoMargins(element, childStyle, width);
                    bool carriesContinuation = ContainsElementOrSelf(element, continuationTarget);
                    List<HtmlFloatExclusion>? childFloats = activeFloats != null
                        && ContainsFloatingDescendant(element, width, childStyle, depth + 1)
                            ? new List<HtmlFloatExclusion>()
                            : null;
                    HtmlRenderFlowBlock LayoutPlacedChild() => LayoutElement(
                        element,
                        width,
                        childStyle,
                        parentStyle,
                        depth + 1,
                        carriesContinuation ? continuationTarget : null,
                        carriesContinuation ? continuationLogicalCharacters : 0,
                        pageBoundary?.Shift(flowHeight),
                        activeFloats?.Count > 0 && !EstablishesFloatContainingBlock(childStyle) ? activeFloats.Select(item => item.Shift(
                            -childStyle.MarginLeft - childStyle.BorderLeftWidth - childStyle.PaddingLeft,
                            -flowHeight - childStyle.MarginTop - childStyle.BorderTopWidth - childStyle.PaddingTop)).ToArray() : null,
                        childFloats, floatContext.At(width, 0D, flowHeight - anticipatedMarginAdjustment));
                    HtmlRenderFlowBlock childBlock = LayoutPlacedChild();
                    if (activeFloats?.Count > 0 && EstablishesFloatContainingBlock(unavoidedStyle)
                        && !unavoidedStyle.ExplicitHeight.HasValue) {
                        // Auto height is only known after layout. Recheck the
                        // float bands with that height so a short box can stay
                        // beside a float without a taller box crossing a later one.
                        for (int attempt = 0; attempt <= activeFloats.Count; attempt++) {
                            double measuredHeight = Math.Max(0.01D,
                                childBlock.Height - childStyle.MarginTop - childStyle.MarginBottom);
                            HtmlRenderBoxStyle measuredStyle = AvoidActiveFloatsForFormattingContext(
                                unavoidedStyle, width, flowHeight, activeFloats, measuredHeight);
                            measuredStyle = ResolveOrdinaryIntrinsicWidths(element, measuredStyle, width, depth + 1);
                            measuredStyle = ResolveNormalFlowHorizontalAutoMargins(element, measuredStyle, width);
                            if (Math.Abs(measuredStyle.MarginTop - childStyle.MarginTop) <= 0.0001D
                                && Math.Abs(measuredStyle.MarginLeft - childStyle.MarginLeft) <= 0.0001D
                                && Math.Abs(measuredStyle.MarginRight - childStyle.MarginRight) <= 0.0001D) break;
                            childStyle = measuredStyle;
                            childFloats?.Clear();
                            childBlock = LayoutPlacedChild();
                        }
                    }
                    if (carriesContinuation) {
                        continuationTarget = null;
                        continuationLogicalCharacters = 0;
                    }
                    double marginAdjustment = 0D;
                    if (childBlock.HasCollapsibleMargins && adjoiningMargins.Count > 0) {
                        adjoiningMargins.Add(childBlock.CollapsibleMarginTopGroup);
                        if (!childBlock.CollapsesThrough) {
                            marginAdjustment = adjoiningMargins.Allocated - adjoiningMargins.Collapsed;
                        }
                    }
                    childBlock = childBlock.AdjustLeadingFlowSpace(marginAdjustment);
                    if (childFloats != null) {
                        foreach (HtmlFloatExclusion exclusion in childFloats) {
                            HtmlFloatExclusion positioned = exclusion.Shift(
                                childStyle.MarginLeft + childStyle.BorderLeftWidth + childStyle.PaddingLeft,
                                flowHeight - marginAdjustment + childStyle.MarginTop + childStyle.BorderTopWidth + childStyle.PaddingTop);
                            activeFloats?.Add(positioned);
                            emittedFloats?.Add(positioned);
                        }
                    }
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
                    } else if (childBlock.CollapsesThrough) {
                        if (adjoiningMargins.Count == 0) {
                            adjoiningMargins.Reset(childBlock.CollapsibleMarginTopGroup);
                        }
                        adjoiningMargins.Add(childBlock.CollapsibleMarginBottomGroup);
                        double collapsed = adjoiningMargins.Collapsed;
                        double trailingAdjustment = adjoiningMargins.Allocated - collapsed;
                        if (Math.Abs(trailingAdjustment) > 0.0001D) {
                            HtmlRenderFlowBlock adjusted = childBlock.AdjustTrailingFlowSpace(trailingAdjustment);
                            blocks[blocks.Count - 1] = adjusted;
                            flowHeight += adjusted.Height - childBlock.Height;
                            childBlock = adjusted;
                        }
                        adjoiningMargins.SetAllocated(collapsed);
                    } else {
                        adjoiningMargins.Reset(childBlock.CollapsibleMarginBottomGroup);
                    }
                    continue;
                }
            }

            inlineNodes.Add(node);
        }

        double trailingInlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth, pageBoundary?.Shift(flowHeight), activeFloats, emittedFloats, flowHeight, floatContext.At(width, 0D, flowHeight));
        flowHeight += trailingInlineHeight;
        if (ownsFloatContext && floatContext.Bottom > flowHeight) blocks.Add(CreateFloatClearanceBlock(width, floatContext.Bottom - flowHeight));
        if (trailingInlineHeight > 0D) adjoiningMargins.Clear();
        if (includeGeneratedAfter) AddGeneratedContentBlock(blocks, container, HtmlPseudoElementKind.After, width, parentStyle);
        return blocks;
    }

    private static IElement? FindDirectChildContaining(IElement container, IElement? target) {
        if (target == null || ReferenceEquals(container, target)) return null;
        IElement? current = target;
        while (current?.ParentElement != null && !ReferenceEquals(current.ParentElement, container)) {
            current = current.ParentElement;
        }
        return current?.ParentElement != null && ReferenceEquals(current.ParentElement, container) ? current : null;
    }

    private static bool ContainsElementOrSelf(IElement candidate, IElement? target) {
        for (IElement? current = target; current != null; current = current.ParentElement) {
            if (ReferenceEquals(candidate, current)) return true;
        }
        return false;
    }

    private struct AdjoiningMarginState {
        private double _positive;
        private double _negative;

        internal int Count { get; private set; }
        internal double Allocated { get; private set; }
        internal double Collapsed => _positive + _negative;

        internal void Add(HtmlCollapsedMargin margin) {
            Count++;
            Allocated += margin.Value;
            _positive = Math.Max(_positive, margin.Positive);
            _negative = Math.Min(_negative, margin.Negative);
        }

        internal void Clear() {
            this = default;
        }

        internal void Reset(HtmlCollapsedMargin margin) {
            this = default;
            Add(margin);
        }

        internal void SetAllocated(double value) {
            Allocated = value;
        }
    }

}
