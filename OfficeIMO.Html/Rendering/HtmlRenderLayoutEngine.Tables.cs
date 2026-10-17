using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock LayoutTable(
        IElement table,
        double containingWidth,
        HtmlRenderBoxStyle style,
        int depth,
        IElement? continuationTarget = null,
        TableFormattingStructure? formatting = null) {
        int legacyBorderWidth = PrepareTableBoxStyleForSizing(table, style, formatting?.Anonymous == true);
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double tableWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, tableWidth - style.HorizontalInsets);
        formatting ??= BuildTableFormattingStructure(table, contentWidth, style, depth);
        string source = HtmlRenderStyleResolver.DescribeSource(table) + (formatting.Anonymous ? ":anonymous-table" : string.Empty);
        string structureKey = GetTableStructureElementKey(formatting.Anonymous ? formatting.Rows.FirstOrDefault()?.Element ?? table : table)
            + (formatting.Anonymous ? ":anonymous-table" : string.Empty);
        _tableFormattingKeys.Add(structureKey);
        HtmlRenderSemanticGroupRole tableRole = formatting.SemanticTable ? HtmlRenderSemanticGroupRole.Table : HtmlRenderSemanticGroupRole.Division;
        ReportUnsupportedTableValues(table, style);
        IReadOnlyList<TableFormattingRow> sourceRows = formatting.Rows;
        if (sourceRows.Count > _options.MaxTableRows) {
            throw new HtmlDomLimitException(
                HtmlRenderDiagnosticCodes.TableLimitExceeded,
                "HTML table row count exceeded the configured maximum.",
                nameof(HtmlRenderOptions.MaxTableRows), sourceRows.Count, _options.MaxTableRows);
        }
        foreach (TableFormattingRow row in sourceRows) ReportTableRowGroupPercentageHeight(row, style);
        IReadOnlyList<TableFormattingRow> headerRows = formatting.HeaderRows;
        IReadOnlyList<TableFormattingRow> footerRows = formatting.FooterRows;
        var headerRowSet = new HashSet<TableFormattingRow>(headerRows);
        var footerRowSet = new HashSet<TableFormattingRow>(footerRows);
        var bodyRows = formatting.BodyRows.ToList();

        // Continuation painting may omit earlier body rows, but their intrinsic
        // contributions still constrain the columns of the same source table.
        IReadOnlyList<TableFormattingRow> sizingRows = formatting.SizingRows;
        TableFormattingRow? continuationRow = continuationTarget == null ? null : sizingRows.FirstOrDefault(row => row.Contains(continuationTarget));
        int skippedBodyRows = 0;
        if (continuationRow != null) {
            int continuationIndex = bodyRows.FindIndex(row => ReferenceEquals(row, continuationRow));
            if (continuationIndex > 0) {
                // An authored table minimum distributes surplus across the full
                // table. Preserve that sizing before removing earlier rows.
                if (style.ExplicitHeight.HasValue) skippedBodyRows = continuationIndex;
                else bodyRows.RemoveRange(0, continuationIndex);
            }
        }
        var rows = new List<TableFormattingRow>(headerRows.Count + bodyRows.Count + footerRows.Count);
        rows.AddRange(headerRows);
        rows.AddRange(bodyRows);
        rows.AddRange(footerRows);
        int rowColumnCount = DetermineColumnCount(sizingRows);
        TableCaptionLayout? caption;
        double topCaptionHeight;
        double bottomCaptionHeight;
        double tableY;
        if (rowColumnCount == 0) {
            ResolveUsedTableHorizontalAutoMargins(table, style, containingWidth, tableWidth, formatting.Anonymous);
            caption = LayoutTableCaption(formatting.Caption, tableWidth, style, depth);
            if (continuationTarget != null && caption != null && caption.Side == "top") caption = null;
            topCaptionHeight = caption != null && caption.Side == "top" ? caption.Height : 0D;
            bottomCaptionHeight = caption != null && caption.Side == "bottom" ? caption.Height : 0D;
            tableY = style.MarginTop + topCaptionHeight;
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.EmptyTable, "A table contained no renderable rows or cells.", HtmlDiagnosticSeverity.Info, source);
            double emptyTableHeight = ResolveTableMinimumHeight(style);
            var emptyVisuals = new List<HtmlRenderVisual>();
            if (caption != null && caption.Side == "top") AppendTableCaption(emptyVisuals, caption, style.MarginLeft, style.MarginTop);
            AddBoxPaint(emptyVisuals, style, style.MarginLeft, tableY, tableWidth, emptyTableHeight, table);
            AddBoxOutlinePaint(emptyVisuals, style, style.MarginLeft, tableY, tableWidth, emptyTableHeight, table);
            if (caption != null && caption.Side == "bottom") AppendTableCaption(emptyVisuals, caption, style.MarginLeft, tableY + emptyTableHeight);
            double emptyHeight = style.MarginTop + topCaptionHeight + emptyTableHeight + bottomCaptionHeight + style.MarginBottom;
            IReadOnlyList<HtmlRenderVisual> semanticEmptyVisuals = new[] {
                new HtmlRenderSemanticGroup(
                    tableRole,
                    style.MarginLeft,
                    style.MarginTop,
                    tableWidth,
                    Math.Max(0.01D, topCaptionHeight + emptyTableHeight + bottomCaptionHeight),
                    emptyVisuals,
                    0,
                    source)
            };
            return new HtmlRenderFlowBlock(style.Display == "inline-table" ? tableWidth + style.MarginLeft + style.MarginRight : containingWidth, emptyHeight, semanticEmptyVisuals, style.BreakBefore, style.BreakAfter, style.AvoidBreakInside, source, pageName: style.PageName);
        }

        int columnCount = Math.Max(rowColumnCount, DetermineDeclaredColumnCount(table));
        CollapsedTableBorderLayout? sourceBorders = ResolveCollapsedTableBorderLayout(table, style, sizingRows,
            columnCount, contentWidth, legacyBorderWidth);
        if (sourceBorders != null) {
            style.TableCollapsedBorderInsets = sourceBorders.TableInsets;
            tableWidth = ResolveBoxWidth(availableWidth, style);
            contentWidth = Math.Max(1D, tableWidth - style.HorizontalInsets);
        }
        IReadOnlyList<TableFormattingRow> fragmentRows = skippedBodyRows > 0
            ? rows.Where((_, index) => index < headerRows.Count || index >= headerRows.Count + skippedBodyRows).ToArray()
            : rows;
        CollapsedTableBorderLayout? fragmentBorders = fragmentRows.SequenceEqual(sizingRows) ? sourceBorders
            : ResolveCollapsedTableBorderLayout(table, style, fragmentRows, columnCount, contentWidth, legacyBorderWidth);
        double horizontalSpacing = style.BorderCollapse == "collapse" ? 0D : style.BorderSpacingX;
        double verticalSpacing = style.BorderCollapse == "collapse" ? 0D : style.BorderSpacingY;
        double trackWidth = Math.Max(0.01D, contentWidth - horizontalSpacing * (columnCount + 1));
        double captionMinimumWidth = Math.Max(0D,
            MeasureTableCaptionMinimumWidth(formatting.Caption, contentWidth, style)
            - style.HorizontalInsets - horizontalSpacing * (columnCount + 1));
        IReadOnlyList<double> columnWidths = ResolveTableColumnWidths(sizingRows, table, columnCount, trackWidth, captionMinimumWidth, style, depth, legacyBorderWidth, sourceBorders, out double usedTrackWidth);
        if ((!style.ExplicitWidth.HasValue && style.TableLayout != "fixed") || usedTrackWidth > trackWidth + 0.0001D) {
            contentWidth = usedTrackWidth + horizontalSpacing * (columnCount + 1);
            tableWidth = contentWidth + style.HorizontalInsets;
        }
        ResolveUsedTableHorizontalAutoMargins(table, style, containingWidth, tableWidth, formatting.Anonymous);
        caption = LayoutTableCaption(formatting.Caption, tableWidth, style, depth);
        if (continuationTarget != null && caption != null && caption.Side == "top") caption = null;
        topCaptionHeight = caption != null && caption.Side == "top" ? caption.Height : 0D;
        bottomCaptionHeight = caption != null && caption.Side == "bottom" ? caption.Height : 0D;
        tableY = style.MarginTop + topCaptionHeight;
        double[] columnOffsets = CreateColumnOffsets(columnWidths);
        var rowLayouts = new List<TableRowLayout>();
        var occupiedColumns = new int[columnCount];
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            CheckCancellation();
            TableFormattingRow row = rows[rowIndex];
            IElement? rowGroupElement = row.GroupElement;
            HtmlRenderBoxStyle? rowGroupStyle = row.GroupStyle;
            HtmlRenderBoxStyle rowStyle = PrepareTablePercentageHeight(row.Style, style);
            var cellLayouts = new List<TableCellLayout>();
            int column = 0;
            double rowHeight = Math.Max(0D, rowStyle.ExplicitHeight ?? 0D);
            foreach (TableFormattingCell cell in row.Cells) {
                int requestedColumnSpan = cell.ColumnSpan;
                column = FindAvailableColumn(occupiedColumns, column, requestedColumnSpan);
                if (column >= columnCount) break;
                int columnSpan = Math.Max(1, Math.Min(requestedColumnSpan, columnCount - column));
                int rowSpan = ReadRowSpan(cell.RowSpan, rows, rowIndex);

                double cellOuterWidth = SumColumnWidths(columnWidths, column, columnSpan) + horizontalSpacing * (columnSpan - 1);
                HtmlRenderBoxStyle cellStyle = ResolveFormattingCellStyle(cell, cellOuterWidth);
                ApplyTableCellFallbackInsets(cellStyle, legacyBorderWidth);
                // Repeated groups change fragment paint topology, not the source cell's
                // authored neighbors or the insets used to size its content.
                if (sourceBorders?.CellInsets.TryGetValue(cell, out HtmlRenderBorderInsets sourceInsets) == true) {
                    cellStyle.TableCollapsedBorderInsets = sourceInsets;
                } else if (fragmentBorders?.CellInsets.TryGetValue(cell, out HtmlRenderBorderInsets collapsedInsets) == true) {
                    cellStyle.TableCollapsedBorderInsets = collapsedInsets;
                }
                if (rowSpan == 1) cellStyle = PrepareTablePercentageHeight(cellStyle, style);
                else if (style.ExplicitHeight.HasValue && cellStyle.TablePercentageHeight.Length > 0) {
                    ReportTablePercentageHeightFallback(cell.Element, cellStyle, "percentage rowspan height");
                }

                double cellContentWidth = Math.Max(1D, cellOuterWidth - cellStyle.HorizontalInsets);
                TableCellPercentageContent? percentageContent = PrepareTableCellPercentageContent(cell, cellContentWidth, cellStyle, style, rowSpan, depth + 1);
                HtmlRenderBoxStyle contentStyle = cellStyle;
                if (percentageContent?.Supported == true) {
                    // Percentage content does not feed its final cell-height
                    // basis back into the row's intrinsic minimum.
                    contentStyle = cellStyle.Clone();
                    contentStyle.ExplicitHeight = null;
                }
                HtmlInlineLayout inline = LayoutTableCellContent(cell, cellContentWidth, contentStyle, depth + 1);
                if (percentageContent != null) percentageContent.LogicalTextOrderCount = _nextLogicalTextOrder - percentageContent.LogicalTextOrderStart;
                double cellHeight = ResolveTableCellMinimumHeight(cellStyle, inline);
                if (rowSpan == 1) rowHeight = Math.Max(rowHeight, cellHeight);
                cellLayouts.Add(new TableCellLayout(cell.Element, cellStyle, inline, column, columnSpan, rowSpan, cellOuterWidth, cellHeight, cell.Nodes != null, cell.StructureKey, percentageContent));
                for (int occupiedColumn = column; occupiedColumn < column + columnSpan; occupiedColumn++) {
                    occupiedColumns[occupiedColumn] = Math.Max(occupiedColumns[occupiedColumn], rowSpan);
                }

                column += columnSpan;
                if (column >= columnCount) break;
            }

            rowLayouts.Add(new TableRowLayout(
                row.Element,
                rowStyle,
                rowGroupElement,
                rowGroupStyle,
                cellLayouts,
                rowHeight,
                headerRowSet.Contains(row),
                footerRowSet.Contains(row), row.Anonymous, row.ContinuationElement, row.StructureKey));
            DecrementOccupancy(occupiedColumns);
        }

        ResolveTableBaselines(rowLayouts);
        ResolveSpanningRowHeights(rowLayouts, verticalSpacing);
        ApplyTableMinimumHeight(rowLayouts, style, verticalSpacing);
        ResolveTableCellPercentageContent(rowLayouts, verticalSpacing, depth + 1);
        ResolveTableCellContentOffsets(rowLayouts, verticalSpacing);
        if (skippedBodyRows > 0) rowLayouts.RemoveRange(headerRows.Count, skippedBodyRows);

        double rowsHeight = verticalSpacing * (rowLayouts.Count + 1);
        for (int rowIndex = 0; rowIndex < rowLayouts.Count; rowIndex++) rowsHeight += rowLayouts[rowIndex].Height;
        double tableHeight = style.VerticalInsets + rowsHeight;
        var visuals = new List<HtmlRenderVisual>();
        var breakOffsets = new List<double>();
        var continuationVisuals = new List<HtmlRenderVisual>();
        var trailingVisuals = new List<HtmlRenderVisual>();
        var runningStringAssignments = new List<HtmlCssRunningStringAssignment>();
        var continuationBreakProgress = new List<HtmlInlineBreakProgress>();
        var forcedBreaks = new List<HtmlRenderForcedBreak>();
        double continuationHeight = 0D;
        double trailingStart = 0D;
        double trailingHeight = 0D;
        var navigationDestinations = new List<HtmlRenderVisual>();
        bool collectingLeadingHeaders = true;
        IReadOnlyList<bool> canBreakAfterRows = ResolveRowBreakAvailability(rowLayouts);
        if (caption != null && caption.Side == "top") AppendTableCaption(visuals, caption, style.MarginLeft, style.MarginTop);
        bool paintSeparateBorders = style.BorderCollapse != "collapse";
        int tablePaintStart = visuals.Count;
        AddBoxPaint(visuals, style, style.MarginLeft, tableY, tableWidth, tableHeight, table, paintSeparateBorders);
        HtmlRenderShape? tableBackdrop = visuals.Skip(tablePaintStart).OfType<HtmlRenderShape>()
            .FirstOrDefault(visual => visual.CanExtendFragmentBackground);
        int rowBackgroundStart = visuals.Count;
        var rowBackgroundLayers = new List<TableRowBackgroundLayer>();
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double rowY = tableY + style.BorderTopWidth + style.PaddingTop + verticalSpacing;
        double headerStart = rowY;
        var hasBodyAfter = new bool[rowLayouts.Count];
        bool bodySeen = false;
        for (int rowIndex = rowLayouts.Count - 1; rowIndex >= 0; rowIndex--) {
            hasBodyAfter[rowIndex] = bodySeen;
            TableRowLayout candidate = rowLayouts[rowIndex];
            if (!candidate.IsHeader && !candidate.IsFooter) bodySeen = true;
        }

        for (int rowIndex = 0; rowIndex < rowLayouts.Count; rowIndex++) {
            CheckCancellation();
            TableRowLayout row = rowLayouts[rowIndex];
            int rowVisualStart = visuals.Count;
            if (row.IsFooter && trailingVisuals.Count == 0) trailingStart = rowY;
            var rowVisuals = new List<HtmlRenderVisual>();
            double rowPaintX = contentX + horizontalSpacing;
            double rowPaintWidth = Math.Max(0.01D, contentWidth - horizontalSpacing * 2D);
            bool startsRowGroup = row.GroupElement != null
                && (rowIndex == 0 || !ReferenceEquals(rowLayouts[rowIndex - 1].GroupElement, row.GroupElement));
            if (startsRowGroup && row.GroupStyle != null) {
                int groupRowCount = 0;
                double groupHeight = 0D;
                for (int groupRowIndex = rowIndex;
                     groupRowIndex < rowLayouts.Count && ReferenceEquals(rowLayouts[groupRowIndex].GroupElement, row.GroupElement);
                     groupRowIndex++) {
                    groupRowCount++;
                    groupHeight += rowLayouts[groupRowIndex].Height;
                }
                groupHeight += verticalSpacing * Math.Max(0, groupRowCount - 1);
                string groupSource = HtmlRenderStyleResolver.DescribeSource(row.GroupElement!);
                var groupBackground = new List<HtmlRenderVisual>();
                AddBoxBackground(groupBackground, row.GroupStyle, rowPaintX, rowY, rowPaintWidth, groupHeight, 0D, row.GroupElement!, groupSource, groupSource);
                if (groupBackground.Count > 0)
                    rowBackgroundLayers.Add(new TableRowBackgroundLayer(groupBackground, true, row.IsHeader, row.IsFooter));
                AddElementNamedDestination(navigationDestinations, row.GroupElement!, rowPaintX, rowY, navigationDestinations.Count);
            }
            string rowSource = HtmlRenderStyleResolver.DescribeSource(row.Element);
            var rowBackground = new List<HtmlRenderVisual>();
            AddBoxBackground(rowBackground, row.Style, rowPaintX, rowY, rowPaintWidth, row.Height, 0D, row.Element, rowSource, rowSource);
            if (rowBackground.Count > 0)
                rowBackgroundLayers.Add(new TableRowBackgroundLayer(rowBackground, false, row.IsHeader, row.IsFooter));
            if (!row.Anonymous) AddElementNamedDestination(navigationDestinations, row.Element, rowPaintX, rowY, navigationDestinations.Count);
            foreach (TableCellLayout cell in row.Cells) {
                double logicalCellX = horizontalSpacing + columnOffsets[cell.Column] + horizontalSpacing * cell.Column;
                double cellX = contentX + (style.Direction == "rtl" ? contentWidth - logicalCellX - cell.Width : logicalCellX);
                double cellHeight = GetSpanningHeight(rowLayouts, rowIndex, cell.RowSpan, verticalSpacing);
                var cellVisuals = new List<HtmlRenderVisual>();
                AddBoxPaint(cellVisuals, cell.Style, cellX, rowY, cell.Width, cellHeight, cell.Element, paintSeparateBorders);
                if (!cell.Anonymous) AddElementNamedDestination(navigationDestinations, cell.Element, cellX, rowY, navigationDestinations.Count);
                double textX = cellX + cell.Style.BorderLeftWidth + cell.Style.PaddingLeft;
                double textY = rowY + cell.ContentOffsetY;
                foreach (HtmlRenderVisual visual in cell.Inline.Visuals) {
                    cellVisuals.Add(visual.Translate(textX, textY, cellVisuals.Count));
                }
                if (!cell.Anonymous) runningStringAssignments.AddRange(ResolveRunningStringAssignments(
                    cell.Element, cell.Style, textY));
                foreach (HtmlCssRunningStringAssignment assignment in cell.Inline.RunningStringAssignments) {
                    runningStringAssignments.Add(assignment.Translate(textY));
                }
                AddBoxOutlinePaint(cellVisuals, cell.Style, cellX, rowY, cell.Width, cellHeight, cell.Element);
                bool headerCell = formatting.SemanticTable && !cell.Anonymous && HasHeaderCellSemantics(cell.Element);
                HtmlRenderSemanticGroupRole cellRole = formatting.SemanticTable && (cell.Anonymous || HasCellSemantics(cell.Element))
                    ? headerCell ? HtmlRenderSemanticGroupRole.TableHeaderCell : HtmlRenderSemanticGroupRole.TableCell
                    : !cell.Anonymous && TryResolveSemanticGroupRole(cell.Element.TagName, out HtmlRenderSemanticGroupRole authoredRole) ? authoredRole : HtmlRenderSemanticGroupRole.Division;
                rowVisuals.Add(new HtmlRenderSemanticGroup(
                    cellRole,
                    cellX,
                    rowY,
                    cell.Width,
                    Math.Max(0.01D, cellHeight),
                    cellVisuals,
                    rowVisuals.Count,
                    HtmlRenderStyleResolver.DescribeSource(cell.Element),
                    cell.Span,
                    cell.RowSpan,
                    headerCell ? ResolveTableHeaderScope(cell.Element) : null,
                    structureElementKey: cell.StructureKey ?? GetTableStructureElementKey(cell.Element)));
            }

            visuals.Add(new HtmlRenderSemanticGroup(
                formatting.SemanticTable && (row.Anonymous || HasRowSemantics(row.Element)) ? HtmlRenderSemanticGroupRole.TableRow : HtmlRenderSemanticGroupRole.Division,
                contentX,
                rowY,
                contentWidth,
                Math.Max(0.01D, row.Height),
                rowVisuals,
                visuals.Count,
                HtmlRenderStyleResolver.DescribeSource(row.Element),
                structureElementKey: row.StructureKey ?? GetTableStructureElementKey(row.Element)));

            bool rowIndependent = !row.IsHeader
                && !row.IsFooter
                && canBreakAfterRows[rowIndex]
                && (rowIndex == 0 || canBreakAfterRows[rowIndex - 1]);
            if (rowIndependent
                && !row.Style.AvoidBreakInside
                && row.GroupStyle?.AvoidBreakInside != true) {
                AddTableRowInternalBreakOffsets(breakOffsets, row, rowY);
            }

            if (collectingLeadingHeaders && row.IsHeader) {
                for (int visualIndex = rowVisualStart; visualIndex < visuals.Count; visualIndex++) {
                    continuationVisuals.Add(visuals[visualIndex].Translate(0D, -headerStart, continuationVisuals.Count));
                }

                continuationHeight += row.Height + verticalSpacing;
            } else {
                collectingLeadingHeaders = false;
            }

            if (row.IsFooter) {
                for (int visualIndex = rowVisualStart; visualIndex < visuals.Count; visualIndex++) {
                    trailingVisuals.Add(visuals[visualIndex].Translate(0D, -trailingStart, trailingVisuals.Count));
                }

                trailingHeight += row.Height + verticalSpacing;
            }

            rowY += row.Height + verticalSpacing;
            HtmlPageBreakTarget forcedTarget = row.Style.BreakAfter;
            if (forcedTarget == HtmlPageBreakTarget.None && rowIndex + 1 < rowLayouts.Count) {
                forcedTarget = rowLayouts[rowIndex + 1].Style.BreakBefore;
            }
            if (forcedTarget != HtmlPageBreakTarget.None && canBreakAfterRows[rowIndex]) {
                forcedBreaks.Add(new HtmlRenderForcedBreak(rowY, forcedTarget));
            }
            bool headerHasBodyAfter = row.IsHeader && hasBodyAfter[rowIndex];
            if (!headerHasBodyAfter && canBreakAfterRows[rowIndex]) {
                breakOffsets.Add(rowY);
                TableRowLayout? nextBodyRow = null;
                for (int nextRowIndex = rowIndex + 1; nextRowIndex < rowLayouts.Count; nextRowIndex++) {
                    TableRowLayout candidate = rowLayouts[nextRowIndex];
                    if (candidate.IsHeader || candidate.IsFooter) continue;
                    nextBodyRow = candidate;
                    break;
                }
                if (nextBodyRow?.ContinuationElement != null) {
                    continuationBreakProgress.Add(new HtmlInlineBreakProgress(rowY, 0, nextBodyRow.ContinuationElement));
                }
            }
        }
        InsertTableRowBackgroundLayers(visuals, rowBackgroundLayers, rowBackgroundStart, 0D);
        InsertTableRowBackgroundLayers(continuationVisuals,
            rowBackgroundLayers.Where(layer => layer.IsHeader), 0, -headerStart);
        InsertTableRowBackgroundLayers(trailingVisuals,
            rowBackgroundLayers.Where(layer => layer.IsFooter), 0, -trailingStart);
        if (style.BorderCollapse == "collapse") {
            int borderVisualStart = visuals.Count;
            AddCollapsedTableBorders(visuals, table, style, fragmentBorders, rowLayouts, columnWidths, columnOffsets, contentX, tableY + style.BorderTopWidth + style.PaddingTop);
            IReadOnlyList<HtmlRenderVisual> collapsedBorders = visuals.Skip(borderVisualStart).ToArray();
            if (continuationVisuals.Count > 0) {
                IReadOnlyList<HtmlRenderVisual> headerBorders = collapsedBorders;
                if (!ReferenceEquals(sourceBorders, fragmentBorders) && sourceBorders != null) {
                    // A resumed body's winning edge belongs only to that fragment.
                    // Future headers retain the original source header's grid rules.
                    var sourceHeaderBorders = new List<HtmlRenderVisual>();
                    AddCollapsedTableBorders(sourceHeaderBorders, table, style, sourceBorders,
                        rowLayouts.TakeWhile(row => row.IsHeader).ToArray(), columnWidths, columnOffsets,
                        contentX, tableY + style.BorderTopWidth + style.PaddingTop);
                    headerBorders = sourceHeaderBorders;
                }
                AppendRepeatedCollapsedBorders(continuationVisuals, headerBorders,
                    headerStart, headerStart + continuationHeight);
            }
            if (trailingVisuals.Count > 0) {
                AppendRepeatedCollapsedBorders(trailingVisuals, collapsedBorders,
                    trailingStart, trailingStart + trailingHeight);
            }
            // A repeated header and the resumed body each retain their source edge.
            // Add the body edge after capturing the reusable header, so it is not
            // incorrectly copied into the next repeated header's paint.
            if (!ReferenceEquals(sourceBorders, fragmentBorders) && sourceBorders != null
                && continuationRow?.Cells.FirstOrDefault() is TableFormattingCell resumedCell
                && sourceBorders.Placements.TryGetValue(resumedCell, out CollapsedTableCellPlacement placement)
                && rowLayouts.TakeWhile(row => row.IsHeader).Any()) {
                double boundaryY = tableY + style.BorderTopWidth + style.PaddingTop
                    + rowLayouts.TakeWhile(row => row.IsHeader).Sum(row => row.Height);
                AddCollapsedTableHorizontalBoundary(visuals, sourceBorders, placement.RowIndex,
                    columnWidths, columnOffsets, contentX, boundaryY, style.Direction == "rtl", source);
            }
        }
        double continuationSourceEnd = headerStart + continuationHeight;
        if (style.BorderCollapse == "collapse") {
            // Repeated row groups start at grid centerlines. Reserve the outside
            // halves too, so their paint and following content stay in the page.
            if (continuationVisuals.Count > 0) {
                for (int index = 0; index < continuationVisuals.Count; index++) {
                    continuationVisuals[index] = continuationVisuals[index].Translate(0D, style.BorderTopWidth, index);
                }
                continuationHeight += style.BorderTopWidth;
            }
            if (trailingVisuals.Count > 0) trailingHeight += style.BorderBottomWidth;
        }
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, tableY, tableWidth, tableHeight, table);
        if (caption != null && caption.Side == "bottom") AppendTableCaption(visuals, caption, style.MarginLeft, tableY + tableHeight);

        double outerHeight = style.MarginTop + topCaptionHeight + tableHeight + bottomCaptionHeight + style.MarginBottom;
        breakOffsets.Add(outerHeight);
        var semanticVisuals = new List<HtmlRenderVisual> {
            new HtmlRenderSemanticGroup(
                tableRole,
                style.MarginLeft,
                style.MarginTop,
                tableWidth,
                Math.Max(0.01D, topCaptionHeight + tableHeight + bottomCaptionHeight),
                visuals,
                0,
                source,
                structureElementKey: structureKey)
        };
        semanticVisuals.AddRange(navigationDestinations);
        PrependRepeatedTableBackdrop(continuationVisuals, tableBackdrop, continuationHeight);
        IReadOnlyList<HtmlRenderVisual> semanticContinuationVisuals = continuationVisuals.Count == 0
            ? Array.Empty<HtmlRenderVisual>()
            : new[] {
                new HtmlRenderSemanticGroup(
                    tableRole,
                    style.MarginLeft,
                    0D,
                    tableWidth,
                    Math.Max(0.01D, continuationHeight),
                    continuationVisuals,
                    0,
                    source,
                    structureElementKey: structureKey)
            };
        IReadOnlyList<HtmlRenderVisual> semanticTrailingVisuals = trailingVisuals.Count == 0
            ? Array.Empty<HtmlRenderVisual>()
            : new[] {
                new HtmlRenderSemanticGroup(
                    tableRole,
                    style.MarginLeft,
                    0D,
                    tableWidth,
                    Math.Max(0.01D, trailingHeight),
                    trailingVisuals,
                    0,
                    source,
                    structureElementKey: structureKey)
            };
        double trailingSourceEnd = caption != null && caption.Side == "bottom"
            ? tableY + tableHeight
            : outerHeight;
        IEnumerable<HtmlRenderTrailingGroup> trailingGroups = trailingVisuals.Count > 0 && trailingHeight > 0D
            ? new[] { new HtmlRenderTrailingGroup(0D, trailingStart, trailingSourceEnd, trailingHeight, semanticTrailingVisuals) }
            : Array.Empty<HtmlRenderTrailingGroup>();
        HtmlPageBreakTarget effectiveBreakBefore = style.BreakBefore;
        if (effectiveBreakBefore == HtmlPageBreakTarget.None && rowLayouts.Count > 0) {
            effectiveBreakBefore = rowLayouts[0].Style.BreakBefore;
        }
        return new HtmlRenderFlowBlock(
            style.Display == "inline-table" ? tableWidth + style.MarginLeft + style.MarginRight : containingWidth,
            outerHeight,
            semanticVisuals,
            effectiveBreakBefore,
            style.BreakAfter,
            style.AvoidBreakInside,
            source,
            breakOffsets,
            trailingGroups: trailingGroups,
            continuationVisuals: semanticContinuationVisuals,
            continuationHeight: continuationHeight,
            continuationStartsAfter: continuationSourceEnd,
            pageName: style.PageName,
            runningStringAssignments: runningStringAssignments,
            inlineBreakProgress: continuationBreakProgress,
            supportsInlineContinuationReflow: continuationBreakProgress.Count > 0,
            forcedBreaks: forcedBreaks);
    }

    /// <summary>Applies table border fallback and removes provisional auto-margin placement before measuring columns.</summary>
    private static int PrepareTableBoxStyleForSizing(IElement table, HtmlRenderBoxStyle style, bool anonymous) {
        style.TableCollapsedBorderInsets = null;
        if (style.BorderCollapse == "collapse") {
            style.PaddingTop = style.PaddingRight = style.PaddingBottom = style.PaddingLeft = 0D;
        }
        int legacyBorderWidth = !anonymous && table.LocalName == "table" ? ReadLegacyTableBorderWidth(table) : 0;
        if (legacyBorderWidth > 0 && !style.BorderDeclared) {
            style.Borders = HtmlRenderBorderEdges.Uniform(legacyBorderWidth, "solid", OfficeColor.FromRgb(128, 128, 128));
        }
        // Replace only previously resolved auto-margin space. The remaining
        // margins still carry authored values and float-avoidance offsets.
        RemoveHorizontalAutoMarginShares(style);
        return legacyBorderWidth;
    }

    private static int ReadLegacyTableBorderWidth(IElement table) {
        string? value = table.GetAttribute("border");
        if (value == null) return 0;
        if (value.Length == 0) return 1;
        return int.TryParse(value, out int width) && width > 0 ? width : 0;
    }

    private static void ApplyTableCellFallbackInsets(HtmlRenderBoxStyle cellStyle, int legacyBorderWidth) {
        if (!cellStyle.HasBorderLayout && !cellStyle.BorderDeclared) {
            cellStyle.Borders = legacyBorderWidth > 0
                ? HtmlRenderBorderEdges.Uniform(1D, "solid", OfficeColor.FromRgb(128, 128, 128))
                : HtmlRenderBorderEdges.Uniform(0D, "none", cellStyle.Color);
        }
    }

    private int DetermineColumnCount(IReadOnlyList<TableFormattingRow> rows) {
        var occupancy = new List<int>();
        int maximum = 0;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            int column = 0;
            foreach (TableFormattingCell cell in rows[rowIndex].Cells) {
                int columnSpan = cell.ColumnSpan;
                column = FindAvailableColumn(occupancy, column, columnSpan);
                long columnEnd = (long)column + columnSpan;
                EnsureTableColumnLimit(columnEnd);
                EnsureOccupancySize(occupancy, (int)columnEnd);
                int rowSpan = ReadRowSpan(cell.RowSpan, rows, rowIndex);
                for (int occupiedColumn = column; occupiedColumn < column + columnSpan; occupiedColumn++) {
                    occupancy[occupiedColumn] = Math.Max(occupancy[occupiedColumn], rowSpan);
                }

                column += columnSpan;
                maximum = Math.Max(maximum, column);
            }

            DecrementOccupancy(occupancy);
        }

        return maximum;
    }

    private void EnsureTableColumnLimit(long count) {
        if (count <= _options.MaxTableColumns) return;
        throw new HtmlDomLimitException(
            HtmlRenderDiagnosticCodes.TableLimitExceeded,
            "HTML table column count exceeded the configured maximum.",
            nameof(HtmlRenderOptions.MaxTableColumns),
            count,
            _options.MaxTableColumns);
    }

    private static int FindAvailableColumn(IReadOnlyList<int> occupancy, int start, int span) {
        int column = Math.Max(0, start);
        while (column < occupancy.Count) {
            bool available = true;
            for (int offset = 0; offset < span && column + offset < occupancy.Count; offset++) {
                if (occupancy[column + offset] <= 0) continue;
                column += offset + 1;
                available = false;
                break;
            }

            if (available) return column;
        }

        return column;
    }

    private static void EnsureOccupancySize(List<int> occupancy, int size) {
        while (occupancy.Count < size) occupancy.Add(0);
    }

    private static void DecrementOccupancy(IList<int> occupancy) {
        for (int column = 0; column < occupancy.Count; column++) {
            if (occupancy[column] > 0) occupancy[column]--;
        }
    }

    private static int ReadRowSpan(string? value, IReadOnlyList<TableFormattingRow> rows, int rowIndex) {
        int maximum = CountRowsRemainingInGroup(rows, rowIndex);
        if (!HtmlIntegerSemantics.TryParseNonNegativeInteger(value, out int requested)) return 1;
        if (requested == 0) return maximum;
        return Math.Max(1, Math.Min(requested, maximum));
    }

    private static int CountRowsRemainingInGroup(IReadOnlyList<TableFormattingRow> rows, int rowIndex) {
        IElement? group = rows[rowIndex].GroupElement;
        int count = 0;
        for (int index = rowIndex; index < rows.Count && ReferenceEquals(rows[index].GroupElement, group); index++) count++;
        return Math.Max(1, count);
    }

    private static void ResolveSpanningRowHeights(IReadOnlyList<TableRowLayout> rows, double verticalSpacing) {
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            foreach (TableCellLayout cell in rows[rowIndex].Cells) {
                if (cell.RowSpan <= 1) continue;
                double currentHeight = GetSpanningHeight(rows, rowIndex, cell.RowSpan, verticalSpacing);
                double deficit = cell.MinimumHeight - currentHeight;
                if (deficit <= 0.0001D) continue;
                double addition = deficit / cell.RowSpan;
                for (int offset = 0; offset < cell.RowSpan && rowIndex + offset < rows.Count; offset++) rows[rowIndex + offset].Height += addition;
            }
        }
    }

    private static double GetSpanningHeight(IReadOnlyList<TableRowLayout> rows, int rowIndex, int rowSpan, double verticalSpacing) {
        double height = 0D;
        int includedRows = 0;
        for (int offset = 0; offset < rowSpan && rowIndex + offset < rows.Count; offset++) {
            height += rows[rowIndex + offset].Height;
            includedRows++;
        }
        if (includedRows > 1) height += verticalSpacing * (includedRows - 1);
        return height;
    }

    private static IReadOnlyList<bool> ResolveRowBreakAvailability(IReadOnlyList<TableRowLayout> rows) {
        var result = new bool[rows.Count];
        int occupiedThrough = -1;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            foreach (TableCellLayout cell in rows[rowIndex].Cells) {
                occupiedThrough = Math.Max(occupiedThrough, rowIndex + cell.RowSpan - 1);
            }

            result[rowIndex] = occupiedThrough <= rowIndex;
        }

        for (int rowIndex = 0; rowIndex + 1 < rows.Count; rowIndex++) {
            TableRowLayout current = rows[rowIndex];
            TableRowLayout next = rows[rowIndex + 1];
            if (current.Style.AvoidBreakAfter || next.Style.AvoidBreakBefore) result[rowIndex] = false;
            if (current.GroupElement != null
                && ReferenceEquals(current.GroupElement, next.GroupElement)
                && current.GroupStyle?.AvoidBreakInside == true) {
                result[rowIndex] = false;
            }
        }

        return result;
    }

    private static void AddTableRowInternalBreakOffsets(ICollection<double> breakOffsets, TableRowLayout row, double rowY) {
        var candidates = new SortedSet<double>();
        foreach (TableCellLayout cell in row.Cells) {
            double contentTop = cell.ContentOffsetY;
            foreach (double offset in cell.Inline.BreakOffsets) {
                // These offsets already end actual laid-out lines. Subtracting
                // computed CSS leading can discard the penultimate line when
                // its selected font has a different used normal advance.
                if (offset < cell.Inline.Height - 0.0001D) candidates.Add(contentTop + offset);
            }
        }

        foreach (double candidate in candidates) {
            bool safeForEveryCell = true;
            foreach (TableCellLayout cell in row.Cells) {
                double contentTop = cell.ContentOffsetY;
                double relative = candidate - contentTop;
                if (relative <= 0.0001D) continue;
                if (relative >= cell.Inline.Height - 0.0001D) continue;
                if (!cell.Inline.BreakOffsets.Any(offset => Math.Abs(offset - relative) <= 0.0001D)) {
                    safeForEveryCell = false;
                    break;
                }
            }
            if (safeForEveryCell && candidate > 0.0001D && candidate < row.Height - 0.0001D) {
                breakOffsets.Add(rowY + candidate);
            }
        }
    }

    private static bool IsTableCell(IElement element) => string.Equals(element.TagName, "td", StringComparison.OrdinalIgnoreCase) || string.Equals(element.TagName, "th", StringComparison.OrdinalIgnoreCase);

    private string GetTableStructureElementKey(IElement element) =>
        "html-element:" + GetSemanticNodeId(element).ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static int ReadSpan(string? value, int maximum) {
        if (!HtmlIntegerSemantics.TryParsePositiveInteger(value, out int span)) span = 1;
        return Math.Max(1, Math.Min(span, maximum));
    }

    private static HtmlRenderTableHeaderScope ResolveTableHeaderScope(IElement cell) {
        string scope = cell.GetAttribute("scope")?.Trim().ToLowerInvariant() ?? string.Empty;
        if (scope == "row" || scope == "rowgroup") return HtmlRenderTableHeaderScope.Row;
        if (scope == "both") return HtmlRenderTableHeaderScope.Both;
        if (scope.Length == 0 && HtmlAccessibilitySemantics.HasRole(cell, "rowheader")) return HtmlRenderTableHeaderScope.Row;
        return HtmlRenderTableHeaderScope.Column;
    }

    private void ReportUnsupportedTableValues(IElement element, HtmlRenderBoxStyle style) {
        var details = new List<string>(2);
        if (style.UnsupportedCaptionSide.Length > 0) details.Add("caption-side=" + style.UnsupportedCaptionSide);
        if (style.UnsupportedTableLayout.Length > 0) details.Add("table-layout=" + style.UnsupportedTableLayout);
        if (style.UnsupportedBorderCollapse.Length > 0) details.Add("border-collapse=" + style.UnsupportedBorderCollapse);
        if (style.UnsupportedBorderSpacing.Length > 0) details.Add("border-spacing=" + style.UnsupportedBorderSpacing);
        if (details.Count == 0) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.TableValueUnsupported,
            "An unsupported table formatting value used its documented fallback.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            string.Join(";", details));
    }

}
