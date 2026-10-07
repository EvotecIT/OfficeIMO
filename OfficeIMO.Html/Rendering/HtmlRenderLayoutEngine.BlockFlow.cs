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
        InlineFloatContext? floatContext = null) =>
        BuildChildBlocks(
            container,
            container.ChildNodes,
            width,
            parentStyle,
            depth,
            includeGeneratedBefore: true,
            includeGeneratedAfter: true,
            continuationTarget,
            continuationLogicalCharacters, floatContext);

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
        InlineFloatContext? floatContext = null) {
        EnsureDepth(depth, container);
        bool ownsFloatContext = floatContext == null;
        floatContext ??= new InlineFloatContext(width);
        var blocks = new List<HtmlRenderFlowBlock>();
        IElement? continuationChild = FindDirectChildContaining(container, continuationTarget);
        bool seekingContinuation = continuationChild != null;
        if (includeGeneratedBefore && !seekingContinuation) AddGeneratedContentBlock(blocks, container, HtmlPseudoElementKind.Before, width, parentStyle);
        double flowHeight = blocks.Sum(block => block.Height);
        var adjoiningMargins = new AdjoiningMarginState();
        var inlineNodes = new List<INode>();
        foreach (INode node in nodes) {
            CheckCancellation();
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
                if (childStyle.Display == "none") {
                    continue;
                }
                if (HtmlCssRunningElementParser.TryParsePosition(childStyle.Position, out string runningElementName)) {
                    double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth, floatContext.At(width, 0D, flowHeight));
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
                    double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth, floatContext.At(width, 0D, flowHeight));
                    flowHeight += inlineHeight;
                    bool carriesContinuation = ContainsElementOrSelf(element, continuationTarget);
                    IReadOnlyList<HtmlRenderFlowBlock> flattenedBlocks = BuildChildBlocks(
                        element,
                        width,
                        childStyle,
                        depth + 1,
                        carriesContinuation ? continuationTarget : null,
                        carriesContinuation ? continuationLogicalCharacters : 0,
                        floatContext.At(width, 0D, flowHeight));
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
                    double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth, floatContext.At(width, 0D, flowHeight));
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

                if (childStyle.FloatSide != "none") {
                    inlineNodes.Add(node);
                    continue;
                }

                if (HtmlRenderStyleResolver.IsBlockElement(element, childStyle)) {
                    double inlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth, floatContext.At(width, 0D, flowHeight));
                    flowHeight += inlineHeight;
                    if (inlineHeight > 0D) {
                        adjoiningMargins.Clear();
                    }
                    double blockY = Math.Max(flowHeight, floatContext.Clearance(childStyle.ClearSide));
                    childStyle = PlaceIndependentBlockBesideFloats(childStyle, width, floatContext, ref blockY);
                    double clearance = blockY - flowHeight;
                    if (clearance > 0D) {
                        blocks.Add(CreateFloatClearanceBlock(width, clearance));
                        flowHeight += clearance;
                        adjoiningMargins.Clear();
                    }
                    var anticipatedMargins = adjoiningMargins;
                    anticipatedMargins.Add(childStyle.MarginTop);
                    double anticipatedMarginAdjustment = adjoiningMargins.Count > 0
                        ? anticipatedMargins.Allocated - anticipatedMargins.Collapsed : 0D;
                    bool carriesContinuation = ContainsElementOrSelf(element, continuationTarget);
                    HtmlRenderFlowBlock childBlock = LayoutElement(
                        element,
                        width,
                        childStyle,
                        parentStyle,
                        depth + 1,
                        carriesContinuation ? continuationTarget : null,
                        carriesContinuation ? continuationLogicalCharacters : 0,
                        floatContext.At(width, 0D, flowHeight - anticipatedMarginAdjustment));
                    if (carriesContinuation) {
                        continuationTarget = null;
                        continuationLogicalCharacters = 0;
                    }
                    double marginAdjustment = 0D;
                    if (childBlock.HasCollapsibleMargins && adjoiningMargins.Count > 0) {
                        adjoiningMargins.Add(childBlock.CollapsibleMarginTop);
                        if (!childBlock.CollapsesThrough) {
                            marginAdjustment = adjoiningMargins.Allocated - adjoiningMargins.Collapsed;
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
                    } else if (childBlock.CollapsesThrough) {
                        if (adjoiningMargins.Count == 0) {
                            adjoiningMargins.Reset(childBlock.CollapsibleMarginTop);
                        }
                        adjoiningMargins.Add(childBlock.CollapsibleMarginBottom);
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
                        adjoiningMargins.Reset(childBlock.CollapsibleMarginBottom);
                    }
                    continue;
                }
            }

            inlineNodes.Add(node);
        }

        double trailingInlineHeight = FlushInlineNodes(blocks, inlineNodes, width, parentStyle, container, depth, floatContext.At(width, 0D, flowHeight));
        flowHeight += trailingInlineHeight;
        if (trailingInlineHeight > 0D) adjoiningMargins.Clear();
        if (ownsFloatContext && floatContext.Bottom > flowHeight) blocks.Add(CreateFloatClearanceBlock(width, floatContext.Bottom - flowHeight));
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

    private static double CollapseVerticalMargins(double first, double second) {
        double positive = Math.Max(0D, Math.Max(first, second));
        double negative = Math.Min(0D, Math.Min(first, second));
        return positive + negative;
    }

    private struct AdjoiningMarginState {
        private double _positive;
        private double _negative;

        internal int Count { get; private set; }
        internal double Allocated { get; private set; }
        internal double Collapsed => _positive + _negative;

        internal void Add(double margin) {
            Count++;
            Allocated += margin;
            _positive = Math.Max(_positive, margin);
            _negative = Math.Min(_negative, margin);
        }

        internal void Clear() {
            this = default;
        }

        internal void Reset(double margin) {
            this = default;
            Add(margin);
        }

        internal void SetAllocated(double value) {
            Allocated = value;
        }
    }

}
