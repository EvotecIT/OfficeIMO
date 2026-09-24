using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool TryLayoutFlexContainer(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        int depth,
        IElement? continuationTarget,
        out HtmlRenderFlowBlock block) {
        block = null!;
        if (style.FlexWrap != "nowrap" && style.FlexWrap != "wrap" && style.FlexWrap != "wrap-reverse") {
            return false;
        }

        bool row = style.FlexDirection == "row" || style.FlexDirection == "row-reverse";
        bool column = style.FlexDirection == "column" || style.FlexDirection == "column-reverse";
        if (!row && !column) return false;
        if (!TryCollectFlexItems(element, containingWidth, style, depth, captureRunningElements: true, out List<FlexItem> items, out List<HtmlCssRunningStringAssignment> runningElementAssignments)) return false;
        if (column) return TryLayoutColumnFlexContainer(element, containingWidth, style, depth, items, runningElementAssignments, out block);

        return TryLayoutRowFlexContainer(element, containingWidth, style, depth, items, runningElementAssignments, continuationTarget, out block);
    }

    private bool TryCollectFlexItems(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        int depth,
        bool captureRunningElements,
        out List<FlexItem> items,
        out List<HtmlCssRunningStringAssignment> runningElementAssignments,
        bool registerPositionedChildren = true) {
        items = new List<FlexItem>();
        runningElementAssignments = new List<HtmlCssRunningStringAssignment>();
        int sourceIndex = 0;
        AddGeneratedFlexItem(element, HtmlPseudoElementKind.Before, containingWidth, style, ref sourceIndex, items);
        foreach (INode node in element.ChildNodes) {
            if (!TryAddFlexNode(node, containingWidth, style, depth + 1, ref sourceIndex, items,
                    captureRunningElements ? runningElementAssignments : null, registerPositionedChildren)) return false;
        }
        AddGeneratedFlexItem(element, HtmlPseudoElementKind.After, containingWidth, style, ref sourceIndex, items);

        return true;
    }

    private bool TryLayoutRowFlexContainer(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        int depth,
        List<FlexItem> items,
        IReadOnlyList<HtmlCssRunningStringAssignment> runningElementAssignments,
        IElement? continuationTarget,
        out HtmlRenderFlowBlock block) {
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        if (style.UnsupportedColumnGap.Length > 0) {
            ReportUnsupportedFlexValue(element, "column-gap=" + style.UnsupportedColumnGap);
        }
        if (style.FlexWrap != "nowrap" && style.UnsupportedRowGap.Length > 0) {
            ReportUnsupportedFlexValue(element, "row-gap=" + style.UnsupportedRowGap);
        }
        List<FlexItem> orderedItems = items.OrderBy(item => item.Style.Order).ThenBy(item => item.SourceIndex).ToList();
        if (continuationTarget != null) {
            int continuationIndex = orderedItems.FindIndex(item =>
                item.Element != null && ContainsElementOrSelf(item.Element, continuationTarget));
            if (continuationIndex >= 0) orderedItems = orderedItems.Skip(continuationIndex).ToList();
        }
        double gap = orderedItems.Count > 1 ? style.ColumnGap : 0D;
        foreach (FlexItem item in orderedItems) {
            item.Basis = ResolveFlexBasis(item, contentWidth);
            item.AutomaticMinimumMainSize = ResolveFlexAutomaticMinimumWidth(item, contentWidth);
        }
        List<FlexLine> lines = CreateFlexLines(orderedItems, style.FlexWrap, contentWidth, gap, vertical: false);
        foreach (FlexLine line in lines) {
            double availableForItems = Math.Max(0D, contentWidth - gap * Math.Max(0, line.Items.Count - 1));
            ResolveFlexMainSizes(line.Items, availableForItems, vertical: false);
            foreach (FlexItem item in line.Items) {
                item.Block = LayoutFlexItem(item, Math.Max(1D, item.MainSize), style, depth + 1);
            }

            line.CrossSize = line.Items.Count == 0 ? 0D : line.Items.Max(item => item.Block!.Height);
        }

        double rowGap = lines.Count > 1 ? style.RowGap : 0D;
        double naturalCrossSize = lines.Sum(line => line.CrossSize) + rowGap * Math.Max(0, lines.Count - 1);
        double crossSize = ResolveFlexCrossSize(style, naturalCrossSize);
        ResolveFlexLineOffsets(lines, style, crossSize, rowGap, HtmlRenderStyleResolver.DescribeSource(element));
        foreach (FlexLine line in lines) {
            StretchFlexItems(line.Items, style, line.CrossSize, depth);
            ResolveFlexMainOffsets(line.Items, style, contentWidth, gap, style.FlexDirection == "row-reverse", vertical: false, source: HtmlRenderStyleResolver.DescribeSource(element));
            foreach (FlexItem item in line.Items) {
                item.CrossOffset = ResolveFlexCrossOffset(item, style, line.CrossSize);
            }
        }

        double boxHeight = ResolveBoxHeight(crossSize, boxWidth, style);
        double outerHeight = Math.Max(0.01D, style.MarginTop + boxHeight + style.MarginBottom);
        var visuals = new List<HtmlRenderVisual>();
        var positionedRunningStringAssignments = new List<HtmlCssRunningStringAssignment>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        AppendLocalPositionedVisuals(
            element,
            Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
            Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
            style.MarginLeft + style.BorderLeftWidth,
            style.MarginTop + style.BorderTopWidth,
            PositionedPaintBand.Negative,
            visuals,
            positionedRunningStringAssignments);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        var itemPaintLayers = new List<FlowPaintLayer>();
        foreach (FlexLine line in lines) {
            foreach (FlexItem item in line.Items) {
                if (item.Element != null) {
                    RecordNormalFlowPlacement(item.Element, element, item.MainOffset, line.CrossOffset + item.CrossOffset, item.Style);
                }
                itemPaintLayers.Add(new FlowPaintLayer(
                    item.Block!,
                    contentX + item.MainOffset,
                    contentY + line.CrossOffset + item.CrossOffset,
                    itemPaintLayers.Count));
            }
        }
        AppendFlowPaintLayers(visuals, itemPaintLayers);
        AppendLocalPositionedVisuals(
            element,
            Math.Max(1D, boxWidth - style.BorderLeftWidth - style.BorderRightWidth),
            Math.Max(0.01D, boxHeight - style.BorderTopWidth - style.BorderBottomWidth),
            style.MarginLeft + style.BorderLeftWidth,
            style.MarginTop + style.BorderTopWidth,
            PositionedPaintBand.NonNegative,
            visuals,
            positionedRunningStringAssignments);
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);

        var atomicVisualBottoms = new Dictionary<HtmlRenderFlowBlock, double>();
        IEnumerable<double> breakOffsets = lines.SelectMany(line => line.Items.SelectMany(item =>
                item.Block!.BreakOffsets.Select(offset => contentY + line.CrossOffset + item.CrossOffset + offset)))
            .Concat(style.FlexWrap == "nowrap"
                ? Array.Empty<double>()
                : lines.Skip(1).Select(line => contentY + line.CrossOffset))
            .Where(offset => lines.All(line => line.Items.All(item => {
                double localOffset = offset - contentY - line.CrossOffset - item.CrossOffset;
                return IsSafeFlexRowBreak(item.Block!, localOffset, atomicVisualBottoms);
            })))
            .Distinct()
            .OrderBy(offset => offset);
        IEnumerable<HtmlRenderLineBreakGroup> lineBreakGroups = lines.SelectMany(line => line.Items.SelectMany(item =>
            item.Block!.LineBreakGroups.Select(group => group.Translate(contentY + line.CrossOffset + item.CrossOffset))));
        IReadOnlyList<HtmlInlineBreakProgress> continuationBreakProgress = style.FlexWrap == "wrap"
            ? lines.Skip(1)
                .Where(line => line.Items.Count > 0 && line.Items[0].Element != null)
                .Select(line => new HtmlInlineBreakProgress(contentY + line.CrossOffset, 0, line.Items[0].Element))
                .ToList()
            : Array.Empty<HtmlInlineBreakProgress>();

        block = new HtmlRenderFlowBlock(
            containingWidth,
            outerHeight,
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            style.AvoidBreakInside,
            HtmlRenderStyleResolver.DescribeSource(element),
            breakOffsets,
            lineBreakGroups: lineBreakGroups,
            pageName: style.PageName,
            runningStringAssignments: NormalizeRunningElementAssignmentOrder(
                PlaceDirectRunningElementAssignments(
                        runningElementAssignments,
                        lines.SelectMany(line => line.Items.Select(item =>
                            new RunningElementFlowAnchor(item.SourceIndex, contentY + line.CrossOffset + item.CrossOffset))),
                        contentY,
                        contentY + crossSize)
                    .Concat(itemPaintLayers.SelectMany(layer =>
                        layer.Block.RunningStringAssignments.Select(assignment => assignment.Translate(layer.Y))))
                    .Concat(positionedRunningStringAssignments),
                outerHeight),
            inlineBreakProgress: continuationBreakProgress,
            supportsInlineContinuationReflow: continuationBreakProgress.Count > 0);
        return true;
    }

    private static bool IsSafeFlexRowBreak(
        HtmlRenderFlowBlock item,
        double offset,
        IDictionary<HtmlRenderFlowBlock, double> atomicVisualBottoms) {
        if (offset <= 0.0001D || offset >= item.Height - 0.0001D) return true;
        if (item.BreakOffsets.Any(candidate => Math.Abs(candidate - offset) <= 0.0001D)) return true;
        // A stretched column can have a large painted but content-free tail. Its background
        // may fragment while the neighboring column supplies the actual page break.
        if (!atomicVisualBottoms.TryGetValue(item, out double bottom)) {
            bottom = LastAtomicFlexVisualBottom(item.Visuals);
            atomicVisualBottoms[item] = bottom;
        }
        return offset >= bottom - 0.0001D;
    }

    private static double LastAtomicFlexVisualBottom(IEnumerable<HtmlRenderVisual> visuals) {
        double bottom = 0D;
        foreach (HtmlRenderVisual visual in visuals) {
            IReadOnlyList<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderLayoutRegion group => group.Visuals,
                HtmlRenderSemanticGroup group => group.Visuals,
                _ => null
            };
            if (children != null) bottom = Math.Max(bottom, LastAtomicFlexVisualBottom(children));
            else if (visual is HtmlRenderText or HtmlRenderImage or HtmlRenderDrawing or HtmlRenderFormField)
                bottom = Math.Max(bottom, visual.LayoutY + visual.LayoutHeight);
        }
        return bottom;
    }

    private double ResolveFlexBasis(FlexItem item, double availableWidth, int intrinsicDepth = 0) {
        HtmlRenderBoxStyle style = item.Style;
        double boxBasis;
        if (style.FlexBasis != "auto") {
            if (TryResolveLength(style.FlexBasis, availableWidth, style.Font.Size, out double parsed)) {
                boxBasis = Math.Max(0D, parsed) + (style.BorderBox ? 0D : style.HorizontalInsets);
            } else {
                ReportUnsupportedFlexValue(item, "flex-basis=" + style.FlexBasis);
                boxBasis = ResolveFlexAutoBoxBasis(item, availableWidth, intrinsicDepth);
            }
        } else {
            boxBasis = ResolveFlexAutoBoxBasis(item, availableWidth, intrinsicDepth);
        }

        return Math.Max(0D, boxBasis + style.MarginLeft + style.MarginRight);
    }

    private double ResolveFlexAutoBoxBasis(FlexItem item, double availableWidth, int intrinsicDepth) {
        HtmlRenderBoxStyle style = item.Style;
        string tag = item.TagName;
        if (IsReplacedImageElementTag(tag) && item.Element != null) return ResolveReplacedImageBoxWidth(item.Element, style);
        if (style.ExplicitWidth.HasValue) {
            return style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
        }

        if (tag == "table") return availableWidth;
        if (item.Element != null
            && style.Display == "flex"
            && (style.FlexDirection == "row" || style.FlexDirection == "row-reverse")) {
            EnsureDepth(intrinsicDepth + 1, item.Element);
            if (TryCollectFlexItems(item.Element, availableWidth, style, intrinsicDepth + 1,
                    captureRunningElements: false, out List<FlexItem> nestedItems, out _, registerPositionedChildren: false)) {
                double nestedWidth = nestedItems.Sum(child => ResolveFlexIntrinsicItemWidth(child, availableWidth, intrinsicDepth + 1))
                    + style.ColumnGap * Math.Max(0, nestedItems.Count - 1);
                return Math.Min(availableWidth, nestedWidth + style.HorizontalInsets);
            }
        }
        IReadOnlyList<GridIntrinsicTextRun> runs = ResolveGridInFlowTextRuns(item, availableWidth);
        double measured = runs.Count == 0 ? 0D : MeasureGridMaxContentRuns(runs);
        if (item.Element != null) {
            foreach (IElement child in item.Element.Children) {
                HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, availableWidth, style);
                if (childStyle.Display == "none" || childStyle.Position == "absolute" || childStyle.Position == "fixed"
                    || childStyle.ExplicitWidthUsesPercentage
                    || !HtmlRenderStyleResolver.IsBlockElement(child, childStyle)) continue;
                // The flattened text runs omit an in-flow block child's own padding.
                measured = Math.Max(measured, ResolveGridMaxContentContribution(new FlexItem(child, childStyle, 0), availableWidth));
            }
        }
        return Math.Min(availableWidth, measured + style.HorizontalInsets);
    }

    private double ResolveFlexIntrinsicItemWidth(FlexItem item, double availableWidth, int intrinsicDepth) {
        if (item.Style.ExplicitWidthUsesPercentage) {
            // A percentage width cannot define the width of its own indefinite flex parent.
            item.Style = item.Style.Clone();
            item.Style.ExplicitWidth = null;
            item.Style.ExplicitWidthUsesPercentage = false;
        }
        double basis = ResolveFlexBasis(item, availableWidth, intrinsicDepth);
        double maxContent = ResolveFlexAutoBoxBasis(item, availableWidth, intrinsicDepth)
            + item.Style.MarginLeft + item.Style.MarginRight;
        item.AutomaticMinimumMainSize = ResolveFlexAutomaticMinimumWidth(item, availableWidth);
        return ClampFlexMainSize(item, Math.Max(basis, maxContent), vertical: false);
    }

    private double ResolveFlexAutomaticMinimumWidth(FlexItem item, double availableWidth) {
        HtmlRenderBoxStyle style = item.Style;
        if (style.MinWidth.HasValue || style.OverflowX is not ("visible" or "clip")) return 0D;

        double minimum;
        if (IsReplacedImageElementTag(item.TagName) && item.Element != null) {
            minimum = ResolveReplacedImageBoxWidth(item.Element, style);
        } else {
            IReadOnlyList<GridIntrinsicTextRun> runs = ResolveGridInFlowTextRuns(item, availableWidth);
            double content = runs.Count == 0 ? 0D : MeasureGridMinContentRuns(runs);
            content = Math.Max(content, ResolveDescendantReplacedGridContribution(item, availableWidth));
            content = Math.Max(content, ResolveDescendantDefiniteFlexWidth(item, availableWidth));
            minimum = content + style.HorizontalInsets;
        }
        if (style.ExplicitWidth.HasValue)
            minimum = Math.Min(minimum, style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        if (style.MaxWidth.HasValue)
            minimum = Math.Min(minimum, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        return Math.Max(0D, minimum + style.MarginLeft + style.MarginRight);
    }

    private double ResolveDescendantDefiniteFlexWidth(FlexItem item, double availableWidth) =>
        item.Element == null ? 0D : ResolveDescendantDefiniteFlexWidth(item.Element, item.Style, availableWidth, 1);

    private double ResolveDescendantDefiniteFlexWidth(IElement parent, HtmlRenderBoxStyle parentStyle, double availableWidth, int depth) {
        double maximum = 0D;
        foreach (IElement child in parent.Children) {
            EnsureDepth(depth, child);
            if (ShouldSkipElement(child)) continue;
            HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, availableWidth, parentStyle);
            if (childStyle.Display == "none" || childStyle.Position == "absolute" || childStyle.Position == "fixed") continue;
            if (IsReplacedImageElement(child)) continue;

            double contribution;
            if (childStyle.ExplicitWidth.HasValue && !childStyle.ExplicitWidthUsesPercentage && childStyle.Display != "inline") {
                contribution = ResolveBoxWidth(availableWidth, childStyle) + childStyle.MarginLeft + childStyle.MarginRight;
            } else {
                double descendant = ResolveDescendantDefiniteFlexWidth(child, childStyle, availableWidth, depth + 1);
                contribution = descendant > 0D ? ResolveGridMeasuredContribution(childStyle, descendant) : 0D;
            }
            maximum = Math.Max(maximum, contribution);
        }
        return maximum;
    }

    private static string CollapseFlexText(string value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        return string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    private static void ResolveFlexMainSizes(IReadOnlyList<FlexItem> items, double availableForItems, bool vertical) {
        if (items.Count == 0) return;
        double basisTotal = items.Sum(item => item.Basis);
        double free = availableForItems - basisTotal;
        bool growing = free > 0.0001D;
        bool shrinking = free < -0.0001D;
        var unfrozen = new HashSet<FlexItem>(items);
        foreach (FlexItem item in items) {
            item.MainSize = ClampFlexMainSize(item, item.Basis, vertical);
            double factor = growing ? item.Style.FlexGrow : item.Style.FlexShrink * item.Basis;
            bool constrainedAgainstGrowth = growing && item.MainSize + 0.0001D < item.Basis;
            bool constrainedAgainstShrink = shrinking && item.MainSize > item.Basis + 0.0001D;
            if (!growing && !shrinking || factor <= 0D || constrainedAgainstGrowth || constrainedAgainstShrink) {
                unfrozen.Remove(item);
            }
        }

        while (unfrozen.Count > 0) {
            double frozenTotal = items.Where(item => !unfrozen.Contains(item)).Sum(item => item.MainSize);
            double unfrozenBasisTotal = unfrozen.Sum(item => item.Basis);
            double remaining = availableForItems - frozenTotal - unfrozenBasisTotal;
            double maximumFactor = unfrozen.Max(item => growing ? item.Style.FlexGrow : item.Style.FlexShrink);
            if (maximumFactor <= 0D || double.IsNaN(maximumFactor) || double.IsInfinity(maximumFactor)) break;
            double factorTotal = unfrozen.Sum(item => {
                double normalized = (growing ? item.Style.FlexGrow : item.Style.FlexShrink) / maximumFactor;
                return growing ? normalized : normalized * item.Basis;
            });
            if (factorTotal <= 0D) break;

            var newlyFrozen = new List<FlexItem>();
            foreach (FlexItem item in unfrozen) {
                double normalized = (growing ? item.Style.FlexGrow : item.Style.FlexShrink) / maximumFactor;
                double factor = growing ? normalized : normalized * item.Basis;
                double proposed = item.Basis + remaining * factor / factorTotal;
                if (double.IsNaN(proposed) || double.IsInfinity(proposed)) proposed = item.Basis;
                double clamped = ClampFlexMainSize(item, proposed, vertical);
                item.MainSize = clamped;
                if (Math.Abs(clamped - proposed) > 0.0001D) newlyFrozen.Add(item);
            }

            if (newlyFrozen.Count == 0) break;
            foreach (FlexItem item in newlyFrozen) unfrozen.Remove(item);
        }
    }

    private static double ClampFlexMainSize(FlexItem item, double value, bool vertical) {
        HtmlRenderBoxStyle style = item.Style;
        double nonContent = vertical
            ? (style.BorderBox ? 0D : style.VerticalInsets) + style.MarginTop + style.MarginBottom
            : (style.BorderBox ? 0D : style.HorizontalInsets) + style.MarginLeft + style.MarginRight;
        double? declaredMinimum = vertical ? style.MinHeight : style.MinWidth;
        double? declaredMaximum = vertical ? style.MaxHeight : style.MaxWidth;
        double minimum = declaredMinimum.HasValue ? declaredMinimum.Value + nonContent
            : vertical ? 0D : item.AutomaticMinimumMainSize;
        double maximum = declaredMaximum.HasValue ? declaredMaximum.Value + nonContent : double.PositiveInfinity;
        return Math.Max(minimum, Math.Min(maximum, Math.Max(0D, value)));
    }

    private static double ResolveFlexCrossSize(HtmlRenderBoxStyle style, double naturalCrossSize) {
        if (!style.ExplicitHeight.HasValue) return naturalCrossSize;
        return style.BorderBox
            ? Math.Max(0D, style.ExplicitHeight.Value - style.VerticalInsets)
            : style.ExplicitHeight.Value;
    }

    private void StretchFlexItems(IReadOnlyList<FlexItem> items, HtmlRenderBoxStyle containerStyle, double crossSize, int depth) {
        foreach (FlexItem item in items) {
            string alignment = ResolveFlexAlignment(item.Style.AlignSelf, containerStyle.AlignItems);
            if (alignment != "stretch" || item.Style.ExplicitHeight.HasValue || item.Style.MarginTopAuto || item.Style.MarginBottomAuto) continue;
            double targetBoxHeight = Math.Max(0.01D, crossSize - item.Style.MarginTop - item.Style.MarginBottom);
            var stretchedStyle = item.Style.Clone();
            stretchedStyle.ExplicitHeight = stretchedStyle.BorderBox
                ? targetBoxHeight
                : Math.Max(0.01D, targetBoxHeight - stretchedStyle.VerticalInsets);
            item.Style = stretchedStyle;
            item.Block = LayoutFlexItem(item, Math.Max(1D, item.MainSize), containerStyle, depth + 1);
        }
    }

    private void ResolveFlexMainOffsets(IReadOnlyList<FlexItem> items, HtmlRenderBoxStyle style, double availableMainSize, double gap, bool reverse, bool vertical, string source) {
        if (items.Count == 0) return;
        double used = items.Sum(item => item.MainSize) + gap * Math.Max(0, items.Count - 1);
        double remaining = Math.Max(0D, availableMainSize - used);
        int autoMarginCount = items.Sum(item => CountFlexMainAutoMargins(item.Style, vertical));
        if (autoMarginCount > 0 && remaining > 0D) {
            double autoMargin = remaining / autoMarginCount;
            double autoCursor = 0D;
            foreach (FlexItem item in items) {
                if (IsFlexMainLeadingMarginAuto(item.Style, vertical, reverse)) autoCursor += autoMargin;
                item.MainOffset = reverse ? availableMainSize - autoCursor - item.MainSize : autoCursor;
                autoCursor += item.MainSize;
                if (IsFlexMainTrailingMarginAuto(item.Style, vertical, reverse)) autoCursor += autoMargin;
                autoCursor += gap;
            }
            return;
        }

        ResolveJustification(style.JustifyContent, items.Count, remaining, gap, reverse, source, out double start, out double between);
        double cursor = start;
        foreach (FlexItem item in items) {
            item.MainOffset = reverse ? availableMainSize - cursor - item.MainSize : cursor;
            cursor += item.MainSize + between;
        }
    }

    private static int CountFlexMainAutoMargins(HtmlRenderBoxStyle style, bool vertical) => vertical
        ? (style.MarginTopAuto ? 1 : 0) + (style.MarginBottomAuto ? 1 : 0)
        : (style.MarginLeftAuto ? 1 : 0) + (style.MarginRightAuto ? 1 : 0);

    private static bool IsFlexMainLeadingMarginAuto(HtmlRenderBoxStyle style, bool vertical, bool reverse) {
        if (vertical) return reverse ? style.MarginBottomAuto : style.MarginTopAuto;
        return reverse ? style.MarginRightAuto : style.MarginLeftAuto;
    }

    private static bool IsFlexMainTrailingMarginAuto(HtmlRenderBoxStyle style, bool vertical, bool reverse) {
        if (vertical) return reverse ? style.MarginTopAuto : style.MarginBottomAuto;
        return reverse ? style.MarginLeftAuto : style.MarginRightAuto;
    }

    private void ResolveJustification(string value, int itemCount, double remaining, double gap, bool reverse, string source, out double start, out double between) {
        string normalized = value == "normal" ? "flex-start" : value;
        start = 0D;
        between = gap;
        switch (normalized) {
            case "flex-start":
                return;
            case "flex-end":
                start = remaining;
                return;
            case "start":
            case "left":
                start = reverse ? remaining : 0D;
                return;
            case "end":
            case "right":
                start = reverse ? 0D : remaining;
                return;
            case "center":
                start = remaining / 2D;
                return;
            case "space-between":
                if (itemCount > 1) between += remaining / (itemCount - 1D);
                return;
            case "space-around":
                if (itemCount > 0) {
                    double spacing = remaining / itemCount;
                    start = spacing / 2D;
                    between += spacing;
                }
                return;
            case "space-evenly":
                double even = remaining / (itemCount + 1D);
                start = even;
                between += even;
                return;
            default:
                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.FlexValueUnsupported, "An unsupported justify-content value used flex-start.", HtmlDiagnosticSeverity.Warning, source, "justify-content=" + value);
                return;
        }
    }

    private double ResolveFlexCrossOffset(FlexItem item, HtmlRenderBoxStyle containerStyle, double crossSize) {
        string alignment = ResolveFlexAlignment(item.Style.AlignSelf, containerStyle.AlignItems);
        double remaining = Math.Max(0D, crossSize - item.Block!.Height);
        if (item.Style.MarginTopAuto || item.Style.MarginBottomAuto) {
            if (item.Style.MarginTopAuto && item.Style.MarginBottomAuto) return remaining / 2D;
            return item.Style.MarginTopAuto ? remaining : 0D;
        }
        bool reverse = containerStyle.FlexWrap == "wrap-reverse";
        if (alignment == "flex-end") return reverse ? 0D : remaining;
        if (alignment == "end") return remaining;
        if (alignment == "center") return remaining / 2D;
        if (alignment == "stretch" || alignment == "flex-start") return reverse ? remaining : 0D;
        if (alignment == "start") return 0D;
        ReportUnsupportedFlexValue(item, "align-self=" + alignment);
        return 0D;
    }

    private static string ResolveFlexAlignment(string alignSelf, string alignItems) {
        string resolved = alignSelf == "auto" ? alignItems : alignSelf;
        return resolved == "normal" ? "stretch" : resolved;
    }

    private void ReportUnsupportedFlexValue(IElement element, string detail) {
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.FlexValueUnsupported, "A flex property value used a deterministic fallback.", HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element), detail);
    }

    private void ReportUnsupportedFlexValue(FlexItem item, string detail) {
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.FlexValueUnsupported, "A flex property value used a deterministic fallback.", HtmlDiagnosticSeverity.Warning, item.Source, detail);
    }

    private sealed class FlexItem {
        internal FlexItem(IElement element, HtmlRenderBoxStyle style, int sourceIndex) {
            Element = element;
            SourceElement = element;
            Style = style;
            SourceIndex = sourceIndex;
            Source = HtmlRenderStyleResolver.DescribeSource(element);
        }

        internal FlexItem(string text, IElement sourceElement, string source, string? link, HtmlRenderBoxStyle style, int sourceIndex, bool paintAnonymousBox) {
            AnonymousText = text;
            SourceElement = sourceElement;
            Source = source;
            Link = link;
            Style = style;
            SourceIndex = sourceIndex;
            PaintAnonymousBox = paintAnonymousBox;
        }

        internal IElement? Element { get; }
        internal IElement SourceElement { get; }
        internal string AnonymousText { get; } = string.Empty;
        internal string TextContent => Element?.TextContent ?? AnonymousText;
        internal string TagName => Element?.TagName.ToLowerInvariant() ?? string.Empty;
        internal string Source { get; }
        internal string? Link { get; }
        internal bool PaintAnonymousBox { get; }
        internal List<FlattenedSemanticPlacement> FlattenedSemanticPlacements { get; } = new List<FlattenedSemanticPlacement>();
        internal HtmlRenderBoxStyle Style { get; set; }
        internal int SourceIndex { get; }
        internal double Basis { get; set; }
        internal double AutomaticMinimumMainSize { get; set; }
        internal double MainSize { get; set; }
        internal double MainOffset { get; set; }
        internal double CrossBasis { get; set; }
        internal double InitialCrossSize { get; set; }
        internal double CrossOffset { get; set; }
        internal bool HasExplicitCrossSize { get; set; }
        internal HtmlRenderFlowBlock? Block { get; set; }
    }
}
