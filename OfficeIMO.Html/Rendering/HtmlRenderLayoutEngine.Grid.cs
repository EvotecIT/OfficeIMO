using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private bool TryLayoutGridContainer(
        IElement element,
        double containingWidth,
        HtmlRenderBoxStyle style,
        int depth,
        out HtmlRenderFlowBlock block) {
        block = null!;
        if (!TryCollectFlexItems(element, containingWidth, style, depth, captureRunningElements: true, out List<FlexItem> formattingItems, out List<HtmlCssRunningStringAssignment> runningElementAssignments)) return false;
        string source = HtmlRenderStyleResolver.DescribeSource(element);
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        if (style.UnsupportedColumnGap.Length > 0) ReportUnsupportedGridValue(source, "column-gap=" + style.UnsupportedColumnGap);
        if (style.UnsupportedRowGap.Length > 0) ReportUnsupportedGridValue(source, "row-gap=" + style.UnsupportedRowGap);

        double? declaredContentHeight = ResolveGridDeclaredContentHeight(style);
        bool usesColumnSubgrid = IsSubgridTrackList(style.GridTemplateColumns)
            && ReferenceEquals(_activeSubgridOwner, element)
            && _activeSubgridColumnSizes != null;
        double inheritedColumnGap = usesColumnSubgrid && _activeSubgridColumnSizes!.Count > 1
            ? style.ColumnGapWasSpecified ? style.ColumnGap : _activeSubgridColumnGap
            : 0D;
        List<GridTrack> columnTracks = usesColumnSubgrid
            ? ResolveColumnSubgridTrackSizes(_activeSubgridColumnSizes!, contentWidth, _activeSubgridColumnGap, inheritedColumnGap, style)
                .Select(size => GridTrack.Fixed(size, "subgrid"))
                .ToList()
            : ParseGridTracks(style.GridTemplateColumns, contentWidth, percentageReferenceIsDefinite: true, style, source, "grid-template-columns");
        bool authoredRowSubgrid = IsSubgridTrackList(style.GridTemplateRows);
        bool isRowSubgridOwner = authoredRowSubgrid && ReferenceEquals(_activeSubgridOwner, element);
        bool usesRowSubgrid = isRowSubgridOwner
            && _activeSubgridRowSizes != null;
        double inheritedRowGap = usesRowSubgrid && _activeSubgridRowSizes!.Count > 1
            ? style.RowGapWasSpecified ? style.RowGap : _activeSubgridRowGap
            : 0D;
        double inheritedRowExtent = usesRowSubgrid
            ? Math.Max(
                0D,
                _activeSubgridRowSizes!.Sum()
                    + _activeSubgridRowGap * Math.Max(0, _activeSubgridRowSizes!.Count - 1)
                    - style.VerticalInsets)
            : 0D;
        List<GridTrack> rowTracks = usesRowSubgrid
            ? ResolveRowSubgridTrackSizes(_activeSubgridRowSizes!, declaredContentHeight ?? inheritedRowExtent, _activeSubgridRowGap, inheritedRowGap, style)
                .Select(size => GridTrack.Fixed(size, "subgrid"))
                .ToList()
            : isRowSubgridOwner
                ? new List<GridTrack>()
                : ParseGridTracks(
                    style.GridTemplateRows,
                    declaredContentHeight ?? 0D,
                    declaredContentHeight.HasValue,
                    style,
                    source,
                    "grid-template-rows");
        IReadOnlyDictionary<string, GridAreaDefinition> areas = ParseGridTemplateAreas(style.GridTemplateAreas, source, out int areaRowCount, out int areaColumnCount);
        IReadOnlyDictionary<string, int> columnLineNames = ParseGridLineNames(
            style.GridTemplateColumns,
            usesColumnSubgrid ? columnTracks.Count + 1 : (int?)null,
            usesColumnSubgrid ? _activeSubgridColumnLineNames : null);
        IReadOnlyDictionary<string, int> rowLineNames = ParseGridLineNames(
            style.GridTemplateRows,
            usesRowSubgrid ? rowTracks.Count + 1 : isRowSubgridOwner ? _activeSubgridRowLineCount : null,
            isRowSubgridOwner ? _activeSubgridRowLineNames : null);
        int explicitColumnCount = Math.Max(1, Math.Max(columnTracks.Count, areaColumnCount));
        int explicitRowCount = Math.Max(1, Math.Max(rowTracks.Count, areaRowCount));
        List<GridItem> items = PlaceGridItems(formattingItems, explicitColumnCount, explicitRowCount, style, source, areas, columnLineNames, rowLineNames, out int columnCount, out int rowCount);
        if (usesColumnSubgrid) {
            ClampSubgridPlacements(items, columnTracks.Count, rows: false);
            columnCount = columnTracks.Count;
        }
        if (usesRowSubgrid) {
            ClampSubgridPlacements(items, rowTracks.Count, rows: true);
            rowCount = rowTracks.Count;
        }
        CollapseEmptyAutoFitColumns(style, items, columnTracks, ref columnCount);
        rowCount = Math.Max(rowCount, Math.Max(1, areaRowCount));
        EnsureGridTrackCount(columnTracks, columnCount, style.GridAutoColumns, contentWidth, percentageReferenceIsDefinite: true, style, source, "grid-auto-columns");
        double columnGap = columnCount > 1
            ? usesColumnSubgrid && !style.ColumnGapWasSpecified ? _activeSubgridColumnGap : style.ColumnGap
            : 0D;
        List<double> columnSizes = ResolveGridTrackSizes(columnTracks, items, contentWidth, columnGap);
        GridAxisLayout columns = ResolveGridAxisLayout(columnTracks, columnSizes, contentWidth, columnGap, style.JustifyContent, source, "justify-content");

        foreach (GridItem item in items) {
            CheckCancellation();
            double cellWidth = columns.SpanSize(item.Column, item.ColumnSpan);
            ApplyInitialGridItemWidth(item, style, cellWidth);
            item.Block = LayoutGridItem(item, Math.Max(1D, cellWidth), style, depth + 1, columns, columnLineNames, rowLineNames: rowLineNames);
        }

        EnsureGridTrackCount(
            rowTracks,
            rowCount,
            style.GridAutoRows,
            declaredContentHeight ?? 0D,
            declaredContentHeight.HasValue,
            style,
            source,
            "grid-auto-rows");
        double rowGap = rowCount > 1
            ? usesRowSubgrid && !style.RowGapWasSpecified ? _activeSubgridRowGap : style.RowGap
            : 0D;
        List<double> rowSizes = ResolveNaturalGridRows(rowTracks, items, rowGap, declaredContentHeight);
        double naturalContentHeight = rowSizes.Sum() + rowGap * Math.Max(0, rowCount - 1);
        double boxHeight = ResolveBoxHeight(naturalContentHeight, boxWidth, style);
        double contentHeight = Math.Max(0D, boxHeight - style.VerticalInsets);
        GridAxisLayout rows = ResolveGridAxisLayout(rowTracks, rowSizes, contentHeight, rowGap, style.AlignContent, source, "align-content");
        RecordGridPositionedContainingRects(
            element,
            style,
            contentWidth,
            contentHeight,
            columns,
            rows,
            areas,
            columnLineNames,
            rowLineNames);

        foreach (GridItem item in items) {
            CheckCancellation();
            double cellWidth = columns.SpanSize(item.Column, item.ColumnSpan);
            double cellHeight = rows.SpanSize(item.Row, item.RowSpan);
            ApplyFinalGridItemSize(item, style, cellWidth, cellHeight);
            item.Block = LayoutGridItem(item, Math.Max(1D, cellWidth), style, depth + 1, columns, columnLineNames, rows, rowLineNames);
            item.OffsetX = ResolveGridHorizontalOffset(item, style, cellWidth);
        }
        ResolveGridVerticalOffsets(items, style, rows);

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
        foreach (GridItem item in items) {
            CheckCancellation();
            double itemX = contentX + columns.Positions[item.Column] + item.OffsetX;
            double itemY = contentY + rows.Positions[item.Row] + item.OffsetY;
            if (item.Item.Element != null) {
                RecordNormalFlowPlacement(
                    item.Item.Element,
                    element,
                    columns.Positions[item.Column] + item.OffsetX,
                    rows.Positions[item.Row] + item.OffsetY,
                    item.Item.Style);
            }
            itemPaintLayers.Add(new FlowPaintLayer(item.Block!, itemX, itemY, itemPaintLayers.Count));
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

        IEnumerable<double> rowBreakOffsets = Enumerable.Range(1, Math.Max(0, rowCount - 1))
            .Where(boundary => !items.Any(item => item.Row < boundary && item.Row + item.RowSpan > boundary))
            .Select(boundary => contentY + rows.Positions[boundary]);
        var rowItemCountDeltas = new int[rowCount + 1];
        foreach (GridItem item in items) {
            rowItemCountDeltas[item.Row]++;
            rowItemCountDeltas[Math.Min(rowCount, item.Row + item.RowSpan)]--;
        }
        var rowItemCounts = new int[rowCount];
        int activeRowItems = 0;
        for (int row = 0; row < rowCount; row++) {
            activeRowItems += rowItemCountDeltas[row];
            rowItemCounts[row] = activeRowItems;
        }
        IEnumerable<double> itemBreakOffsets = items
            .Where(item => item.RowSpan == 1 && rowItemCounts[item.Row] == 1)
            .SelectMany(item => item.Block!.BreakOffsets.Select(offset =>
                contentY + rows.Positions[item.Row] + item.OffsetY + offset));
        IEnumerable<double> breakOffsets = rowBreakOffsets.Concat(itemBreakOffsets).Distinct().OrderBy(offset => offset);
        block = new HtmlRenderFlowBlock(
            containingWidth,
            outerHeight,
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            style.AvoidBreakInside,
            source,
            breakOffsets,
            pageName: style.PageName,
            runningStringAssignments: NormalizeRunningElementAssignmentOrder(
                PlaceDirectRunningElementAssignments(
                        runningElementAssignments,
                        items.Select(item => new RunningElementFlowAnchor(
                            item.Item.SourceIndex,
                            contentY + rows.Positions[item.Row] + item.OffsetY)),
                        contentY,
                        contentY + contentHeight)
                    .Concat(itemPaintLayers.SelectMany(layer =>
                        layer.Block.RunningStringAssignments.Select(assignment => assignment.Translate(layer.Y))))
                    .Concat(positionedRunningStringAssignments),
                outerHeight));
        return true;
    }

    private static void ClampSubgridPlacements(IReadOnlyList<GridItem> items, int trackCount, bool rows) {
        if (trackCount <= 0) return;
        foreach (GridItem item in items) {
            int start = rows ? item.Row : item.Column;
            int span = rows ? item.RowSpan : item.ColumnSpan;
            int end = Math.Min(trackCount, Math.Max(1, start + span));
            start = Math.Min(trackCount - 1, Math.Max(0, start));
            if (end <= start) end = start + 1;
            if (rows) {
                item.Row = start;
                item.RowSpan = end - start;
            } else {
                item.Column = start;
                item.ColumnSpan = end - start;
            }
        }
    }

    private static IReadOnlyList<double> ResolveColumnSubgridTrackSizes(
        IReadOnlyList<double> inheritedSizes,
        double contentWidth,
        double parentGap,
        double subgridGap,
        HtmlRenderBoxStyle style) {
        var sizes = inheritedSizes.Select(size => Math.Max(0D, size)).ToList();
        if (sizes.Count == 0) return sizes;

        if (sizes.Count == 1) {
            sizes[0] -= style.BorderLeftWidth + style.PaddingLeft + style.BorderRightWidth + style.PaddingRight;
        } else {
            double halfGapDifference = (parentGap - subgridGap) / 2D;
            sizes[0] += halfGapDifference - style.BorderLeftWidth - style.PaddingLeft;
            sizes[sizes.Count - 1] += halfGapDifference - style.BorderRightWidth - style.PaddingRight;
            for (int index = 1; index < sizes.Count - 1; index++) sizes[index] += halfGapDifference * 2D;
        }
        for (int index = 0; index < sizes.Count; index++) sizes[index] = Math.Max(0D, sizes[index]);

        double targetTrackWidth = Math.Max(0D, contentWidth - subgridGap * Math.Max(0, sizes.Count - 1));
        double adjustment = targetTrackWidth - sizes.Sum();
        if (adjustment > 0.000001D) {
            sizes[0] += adjustment / 2D;
            sizes[sizes.Count - 1] += adjustment - adjustment / 2D;
            return sizes;
        }

        double remaining = -adjustment;
        for (int offset = 0; remaining > 0.000001D && offset < sizes.Count; offset++) {
            int left = offset;
            int right = sizes.Count - 1 - offset;
            remaining = ReduceSubgridEdgeTrack(sizes, left, remaining);
            if (remaining > 0.000001D && right != left) remaining = ReduceSubgridEdgeTrack(sizes, right, remaining);
        }
        return sizes;
    }

    private static double ReduceSubgridEdgeTrack(IList<double> sizes, int index, double requested) {
        double applied = Math.Min(Math.Max(0D, requested), sizes[index]);
        sizes[index] -= applied;
        return requested - applied;
    }

    private static IReadOnlyList<double> ResolveRowSubgridTrackSizes(
        IReadOnlyList<double> inheritedSizes,
        double contentHeight,
        double parentGap,
        double subgridGap,
        HtmlRenderBoxStyle style) {
        var sizes = inheritedSizes.Select(size => Math.Max(0D, size)).ToList();
        if (sizes.Count == 0) return sizes;
        if (sizes.Count == 1) {
            sizes[0] -= style.BorderTopWidth + style.PaddingTop + style.BorderBottomWidth + style.PaddingBottom;
        } else {
            double halfGapDifference = (parentGap - subgridGap) / 2D;
            sizes[0] += halfGapDifference - style.BorderTopWidth - style.PaddingTop;
            sizes[sizes.Count - 1] += halfGapDifference - style.BorderBottomWidth - style.PaddingBottom;
            for (int index = 1; index < sizes.Count - 1; index++) sizes[index] += halfGapDifference * 2D;
        }
        for (int index = 0; index < sizes.Count; index++) sizes[index] = Math.Max(0D, sizes[index]);

        double targetTrackHeight = Math.Max(0D, contentHeight - subgridGap * Math.Max(0, sizes.Count - 1));
        double adjustment = targetTrackHeight - sizes.Sum();
        if (adjustment > 0.000001D) {
            sizes[0] += adjustment / 2D;
            sizes[sizes.Count - 1] += adjustment - adjustment / 2D;
            return sizes;
        }
        double remaining = -adjustment;
        for (int offset = 0; remaining > 0.000001D && offset < sizes.Count; offset++) {
            int top = offset;
            int bottom = sizes.Count - 1 - offset;
            remaining = ReduceSubgridEdgeTrack(sizes, top, remaining);
            if (remaining > 0.000001D && bottom != top) remaining = ReduceSubgridEdgeTrack(sizes, bottom, remaining);
        }
        return sizes;
    }

    private HtmlRenderFlowBlock LayoutGridItem(
        GridItem item,
        double containingWidth,
        HtmlRenderBoxStyle parentStyle,
        int depth,
        GridAxisLayout columns,
        IReadOnlyDictionary<string, int> columnLineNames,
        GridAxisLayout? rows = null,
        IReadOnlyDictionary<string, int>? rowLineNames = null) {
        IElement? previousOwner = _activeSubgridOwner;
        IReadOnlyList<double>? previousSizes = _activeSubgridColumnSizes;
        IReadOnlyDictionary<string, int>? previousLineNames = _activeSubgridColumnLineNames;
        double previousGap = _activeSubgridColumnGap;
        IReadOnlyList<double>? previousRowSizes = _activeSubgridRowSizes;
        IReadOnlyDictionary<string, int>? previousRowLineNames = _activeSubgridRowLineNames;
        int? previousRowLineCount = _activeSubgridRowLineCount;
        double previousRowGap = _activeSubgridRowGap;
        try {
            _activeSubgridOwner = item.Item.Element;
            _activeSubgridColumnSizes = columns.Sizes.Skip(item.Column).Take(item.ColumnSpan).ToList();
            _activeSubgridColumnLineNames = columnLineNames
                .Where(pair => pair.Value >= item.Column && pair.Value <= item.Column + item.ColumnSpan)
                .ToDictionary(pair => pair.Key, pair => pair.Value - item.Column, StringComparer.Ordinal);
            _activeSubgridColumnGap = columns.Between;
            _activeSubgridRowSizes = rows?.Sizes.Skip(item.Row).Take(item.RowSpan).ToList();
            _activeSubgridRowLineNames = rowLineNames == null
                ? null
                : rowLineNames.Where(pair => pair.Value >= item.Row && pair.Value <= item.Row + item.RowSpan)
                    .ToDictionary(pair => pair.Key, pair => pair.Value - item.Row, StringComparer.Ordinal);
            _activeSubgridRowLineCount = item.RowSpan + 1;
            _activeSubgridRowGap = rows?.Between ?? 0D;
            return LayoutFlexItem(item.Item, containingWidth, parentStyle, depth);
        } finally {
            _activeSubgridOwner = previousOwner;
            _activeSubgridColumnSizes = previousSizes;
            _activeSubgridColumnLineNames = previousLineNames;
            _activeSubgridColumnGap = previousGap;
            _activeSubgridRowSizes = previousRowSizes;
            _activeSubgridRowLineNames = previousRowLineNames;
            _activeSubgridRowLineCount = previousRowLineCount;
            _activeSubgridRowGap = previousRowGap;
        }
    }

    private double? ResolveGridDeclaredContentHeight(HtmlRenderBoxStyle style) {
        if (!style.ExplicitHeight.HasValue) return null;
        return style.BorderBox
            ? Math.Max(0D, style.ExplicitHeight.Value - style.VerticalInsets)
            : Math.Max(0D, style.ExplicitHeight.Value);
    }

    private static List<double> ResolveNaturalGridRows(
        IReadOnlyList<GridTrack> tracks,
        IReadOnlyList<GridItem> items,
        double gap,
        double? declaredContentHeight) {
        var sizes = tracks.Select(track => track.Kind == GridTrackKind.Fixed ? Math.Max(track.Value, track.Minimum) : track.Minimum).ToList();
        foreach (GridItem item in items.OrderBy(item => item.RowSpan)) {
            double required = item.Block!.Height;
            double current = sizes.Skip(item.Row).Take(item.RowSpan).Sum() + gap * Math.Max(0, item.RowSpan - 1);
            double deficit = Math.Max(0D, required - current);
            if (deficit <= 0D) continue;
            List<int> flexible = Enumerable.Range(item.Row, item.RowSpan).Where(index => tracks[index].Kind != GridTrackKind.Fixed).ToList();
            if (flexible.Count == 0) flexible.AddRange(Enumerable.Range(item.Row, item.RowSpan));
            double addition = deficit / flexible.Count;
            foreach (int index in flexible) sizes[index] += addition;
        }

        if (declaredContentHeight.HasValue) {
            double trackSpace = Math.Max(0D, declaredContentHeight.Value - gap * Math.Max(0, tracks.Count - 1));
            double fractions = tracks.Where(track => track.Kind == GridTrackKind.Fraction).Sum(track => track.Value);
            if (fractions > 0D) DistributeGridFractions(tracks, sizes, trackSpace);
        }
        return sizes;
    }

    private void ApplyInitialGridItemWidth(GridItem item, HtmlRenderBoxStyle containerStyle, double cellWidth) {
        string alignment = ResolveGridAlignment(item.Item.Style.JustifySelf, containerStyle.JustifyItems);
        double targetWidth = alignment == "stretch" && !item.HasExplicitWidth && !HasHorizontalAutoMargin(item.Item.Style)
            ? cellWidth
            : ResolveColumnFlexCrossBasis(item.Item, cellWidth);
        ApplyGridItemWidth(item, targetWidth);
    }

    private void ApplyFinalGridItemSize(GridItem item, HtmlRenderBoxStyle containerStyle, double cellWidth, double cellHeight) {
        HtmlRenderBoxStyle style = item.Item.Style.Clone();
        string horizontal = ResolveGridAlignment(style.JustifySelf, containerStyle.JustifyItems);
        string vertical = ResolveGridAlignment(style.AlignSelf, containerStyle.AlignItems);
        if (horizontal == "stretch" && !item.HasExplicitWidth && !HasHorizontalAutoMargin(style)) {
            double targetBoxWidth = Math.Max(0.01D, cellWidth - style.MarginLeft - style.MarginRight);
            style.ExplicitWidth = style.BorderBox ? targetBoxWidth : Math.Max(0.01D, targetBoxWidth - style.HorizontalInsets);
        }
        if (vertical == "stretch" && !item.HasExplicitHeight && !HasVerticalAutoMargin(style)) {
            double targetBoxHeight = Math.Max(0.01D, cellHeight - style.MarginTop - style.MarginBottom);
            style.ExplicitHeight = style.BorderBox ? targetBoxHeight : Math.Max(0.01D, targetBoxHeight - style.VerticalInsets);
        }
        item.Item.Style = style;
    }

    private static void ApplyGridItemWidth(GridItem item, double targetOuterWidth) {
        if (item.HasExplicitWidth) return;
        HtmlRenderBoxStyle style = item.Item.Style.Clone();
        double targetBoxWidth = Math.Max(0.01D, targetOuterWidth - style.MarginLeft - style.MarginRight);
        style.ExplicitWidth = style.BorderBox ? targetBoxWidth : Math.Max(0.01D, targetBoxWidth - style.HorizontalInsets);
        item.Item.Style = style;
    }

    private double ResolveGridHorizontalOffset(GridItem item, HtmlRenderBoxStyle containerStyle, double cellWidth) {
        HtmlRenderBoxStyle style = item.Item.Style;
        double outerWidth = ResolveColumnFlexOuterWidth(style, cellWidth);
        double remaining = Math.Max(0D, cellWidth - outerWidth);
        if (style.MarginLeftAuto || style.MarginRightAuto) {
            if (style.MarginLeftAuto && style.MarginRightAuto) return remaining / 2D;
            return style.MarginLeftAuto ? remaining : 0D;
        }
        return ResolveGridAlignmentOffset(ResolveGridAlignment(style.JustifySelf, containerStyle.JustifyItems), remaining, item.Item.Source, "justify-self");
    }

    private double ResolveGridVerticalOffset(GridItem item, HtmlRenderBoxStyle containerStyle, double cellHeight) {
        HtmlRenderBoxStyle style = item.Item.Style;
        double remaining = Math.Max(0D, cellHeight - item.Block!.Height);
        if (style.MarginTopAuto || style.MarginBottomAuto) {
            if (style.MarginTopAuto && style.MarginBottomAuto) return remaining / 2D;
            return style.MarginTopAuto ? remaining : 0D;
        }
        return ResolveGridAlignmentOffset(ResolveGridAlignment(style.AlignSelf, containerStyle.AlignItems), remaining, item.Item.Source, "align-self");
    }

    private void ResolveGridVerticalOffsets(
        IReadOnlyList<GridItem> items,
        HtmlRenderBoxStyle containerStyle,
        GridAxisLayout rows) {
        foreach (GridItem item in items) {
            double cellHeight = rows.SpanSize(item.Row, item.RowSpan);
            item.OffsetY = ResolveGridVerticalOffset(item, containerStyle, cellHeight);
        }

        foreach (IGrouping<int, GridItem> rowGroup in items
            .Where(item => item.RowSpan == 1
                && !HasVerticalAutoMargin(item.Item.Style)
                && ResolveGridAlignment(item.Item.Style.AlignSelf, containerStyle.AlignItems) == "baseline")
            .GroupBy(item => item.Row)) {
            double sharedBaseline = rowGroup.Max(ResolveGridItemBaseline);
            foreach (GridItem item in rowGroup) {
                double remaining = Math.Max(0D, rows.SpanSize(item.Row, 1) - item.Block!.Height);
                item.OffsetY = Math.Min(remaining, Math.Max(0D, sharedBaseline - ResolveGridItemBaseline(item)));
            }
        }
    }

    private static double ResolveGridItemBaseline(GridItem item) {
        HtmlRenderText? firstText = EnumerateGridTextVisuals(item.Block!.Visuals)
            .OrderBy(text => text.LayoutY)
            .ThenBy(text => text.X)
            .FirstOrDefault();
        if (firstText == null) return item.Block.Height;
        double leading = Math.Max(0D, firstText.LineHeight - firstText.Font.Size);
        return firstText.LayoutY + Math.Min(firstText.LineHeight, leading / 2D + firstText.Font.Size * 0.8D);
    }

    private static IEnumerable<HtmlRenderText> EnumerateGridTextVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            if (visual is HtmlRenderText text) yield return text;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderClipGroup clipGroup
                ? clipGroup.Visuals
                : visual is HtmlRenderPathClipGroup pathClipGroup
                    ? pathClipGroup.Visuals
                    : visual is HtmlRenderEffectGroup effectGroup
                        ? effectGroup.Visuals
                        : visual is HtmlRenderSemanticGroup semanticGroup
                            ? semanticGroup.Visuals
                            : visual is HtmlRenderLogicalTextGroup logicalTextGroup
                                ? logicalTextGroup.Visuals
                                : visual is HtmlRenderFormField formField ? formField.Visuals : null;
            if (children == null) continue;
            foreach (HtmlRenderText child in EnumerateGridTextVisuals(children)) yield return child;
        }
    }

    private double ResolveGridAlignmentOffset(string alignment, double remaining, string source, string property) {
        if (alignment == "end" || alignment == "flex-end") return remaining;
        if (alignment == "center") return remaining / 2D;
        if (alignment == "stretch" || alignment == "start" || alignment == "flex-start" || alignment == "baseline") return 0D;
        ReportUnsupportedGridValue(source, property + "=" + alignment);
        return 0D;
    }

    private static string ResolveGridAlignment(string self, string container) {
        string resolved = self == "auto" ? container : self;
        if (resolved == "normal") return "stretch";
        return resolved == "first baseline" ? "baseline" : resolved;
    }

    private static bool HasHorizontalAutoMargin(HtmlRenderBoxStyle style) => style.MarginLeftAuto || style.MarginRightAuto;
    private static bool HasVerticalAutoMargin(HtmlRenderBoxStyle style) => style.MarginTopAuto || style.MarginBottomAuto;

    private static bool IsSubgridTrackList(string value) {
        string normalized = value?.Trim() ?? string.Empty;
        if (!normalized.StartsWith("subgrid", StringComparison.OrdinalIgnoreCase)) return false;
        int cursor = "subgrid".Length;
        if (cursor < normalized.Length && !char.IsWhiteSpace(normalized[cursor]) && normalized[cursor] != '[') return false;
        while (cursor < normalized.Length) {
            while (cursor < normalized.Length && char.IsWhiteSpace(normalized[cursor])) cursor++;
            if (cursor >= normalized.Length) return true;
            if (normalized[cursor] != '[') return false;
            int close = normalized.IndexOf(']', cursor + 1);
            if (close < 0 || normalized.Substring(cursor + 1, close - cursor - 1).Trim().Length == 0) return false;
            cursor = close + 1;
        }
        return true;
    }
}
