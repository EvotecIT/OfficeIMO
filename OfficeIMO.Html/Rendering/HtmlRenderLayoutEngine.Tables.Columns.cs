using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<double> ResolveTableColumnWidths(IReadOnlyList<TableFormattingRow> rows, IElement table, int columnCount, double contentWidth, double captionMinimumWidth, HtmlRenderBoxStyle tableStyle, int depth, int legacyBorderWidth, out double usedWidth) {
        if (tableStyle.TableLayout == "fixed") {
            var fixedWidths = new double[columnCount];
            ApplyDeclaredColumnWidths(table, contentWidth, tableStyle, fixedWidths, fixedWidths);
            ApplyFirstRowAuthoredWidths(rows, tableStyle, fixedWidths, contentWidth);
            usedWidth = contentWidth;
            return AllocateFixedColumnWidths(fixedWidths, contentWidth);
        }

        ResolveAutoTableColumnContributions(rows, table, columnCount, contentWidth, tableStyle, depth, legacyBorderWidth,
            out double[] minimums, out double[] preferred, out bool[] authoredWidths);
        double minimumAllowed = tableStyle.MinWidth.HasValue
            ? Math.Max(0.01D, tableStyle.MinWidth.Value - (tableStyle.BorderBox ? tableStyle.HorizontalInsets : 0D))
            : 0.01D;
        minimumAllowed = Math.Max(minimumAllowed, captionMinimumWidth);
        usedWidth = tableStyle.ExplicitWidth.HasValue
            ? contentWidth
            : Math.Min(contentWidth, Math.Max(minimumAllowed, preferred.Sum()));
        return AllocateAutoColumnWidths(minimums, preferred, authoredWidths, usedWidth);
    }

    /// <summary>Collects column constraints once for both table layout and non-paint intrinsic sizing.</summary>
    private void ResolveAutoTableColumnContributions(IReadOnlyList<TableFormattingRow> rows, IElement table,
        int columnCount, double contentWidth, HtmlRenderBoxStyle tableStyle, int depth, int legacyBorderWidth,
        out double[] minimums, out double[] preferred, out bool[] authoredWidths) {
        minimums = Enumerable.Repeat(1D, columnCount).ToArray();
        preferred = Enumerable.Repeat(1D, columnCount).ToArray();
        authoredWidths = new bool[columnCount];
        ApplyDeclaredColumnWidths(table, contentWidth, tableStyle, minimums, preferred, authoredWidths);
        var spanningCells = new List<(int Column, int Span, double Minimum, double Preferred)>();
        var occupancy = new int[columnCount];
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            int column = 0;
            foreach (TableFormattingCell cell in rows[rowIndex].Cells) {
                int requestedSpan = Math.Min(cell.ColumnSpan, columnCount);
                column = FindAvailableColumn(occupancy, column, requestedSpan);
                if (column >= columnCount) break;
                int span = Math.Max(1, Math.Min(requestedSpan, columnCount - column));
                HtmlRenderBoxStyle cellStyle = ResolveFormattingCellStyle(cell, contentWidth);
                ApplyTableCellFallbackInsets(cellStyle, legacyBorderWidth);
                if (span == 1 && cellStyle.ExplicitWidth.HasValue && !cellStyle.ExplicitWidthUsesPercentage) {
                    authoredWidths[column] = true;
                }
                ResolveTableCellIntrinsicWidths(cell, cellStyle, contentWidth, depth, out double minimum, out double maximum);
                if (span == 1) {
                    minimums[column] = Math.Max(minimums[column], minimum);
                    preferred[column] = Math.Max(preferred[column], maximum);
                } else {
                    spanningCells.Add((column, span, minimum, maximum));
                }
                int rowSpan = ReadRowSpan(cell.RowSpan, rows, rowIndex);
                for (int occupied = column; occupied < column + span; occupied++) occupancy[occupied] = Math.Max(occupancy[occupied], rowSpan);
                column += span;
            }
            DecrementOccupancy(occupancy);
        }
        ApplyTableSpanningWidths(minimums, preferred, spanningCells,
            tableStyle.BorderCollapse == "collapse" ? 0D : tableStyle.BorderSpacingX);
    }

    /// <summary>Measures a table as one atomic box using source columns, without painting or feeding layout widths back into measurement.</summary>
    private (double Minimum, double Maximum) ResolveIntrinsicTableOuterWidths(IElement table,
        HtmlRenderBoxStyle style, double availableWidth, int depth) {
        EnsureDepth(depth, table);
        style = _styleResolver.ResolveIntrinsicMeasurementStyle(table, style);
        if (!style.HasIntrinsicWidths && style.ExplicitWidth.HasValue && !style.ExplicitWidthUsesPercentage
            && TryResolveDefiniteGridContribution(new FlexItem(table, style, 0), availableWidth, out double definite)) return (definite, definite);
        TableFormattingStructure formatting = BuildTableFormattingStructure(table, availableWidth, style, depth);
        int columnCount = Math.Max(DetermineColumnCount(formatting.Rows), DetermineDeclaredColumnCount(table));
        double captionMinimum = Math.Max(0D, MeasureTableCaptionMinimumWidth(formatting.Caption, availableWidth, style) - style.HorizontalInsets);
        double minimum = captionMinimum;
        double maximum = captionMinimum;
        if (columnCount > 0) {
            int legacyBorderWidth = table.LocalName == "table" ? ReadLegacyTableBorderWidth(table) : 0;
            ResolveAutoTableColumnContributions(formatting.Rows, table, columnCount, availableWidth, style, depth, legacyBorderWidth,
                out double[] minimums, out double[] preferred, out _);
            double spacing = style.BorderCollapse == "collapse" ? 0D : style.BorderSpacingX * (columnCount + 1);
            minimum = Math.Max(minimum, minimums.Sum() + spacing);
            maximum = Math.Max(maximum, preferred.Sum() + spacing);
        }
        return style.HasIntrinsicWidths
            ? ResolveOrdinaryIntrinsicContributions(style, minimum, maximum, availableWidth)
            : (ResolveGridMeasuredContribution(style, minimum), ResolveGridMeasuredContribution(style, maximum));
    }

    private double MeasureTableCaptionMinimumWidth(IElement? caption, double containingWidth, HtmlRenderBoxStyle tableStyle) {
        if (caption == null) return 0D;
        HtmlRenderBoxStyle captionStyle = _styleResolver.Resolve(caption, containingWidth, tableStyle);
        if (captionStyle.Display == "none") return 0D;
        IReadOnlyList<IntrinsicTextRun> runs = ResolveInFlowIntrinsicTextRuns(
            new FlexItem(caption, captionStyle, 0), containingWidth, includeDescendantInsets: true);
        double minimum = MeasureMinContentRuns(runs) + captionStyle.HorizontalInsets;
        if (captionStyle.ExplicitWidth.HasValue && !captionStyle.ExplicitWidthUsesPercentage) {
            minimum = Math.Max(minimum, captionStyle.ExplicitWidth.Value
                + (captionStyle.BorderBox ? 0D : captionStyle.HorizontalInsets));
        }
        return minimum + captionStyle.MarginLeft + captionStyle.MarginRight;
    }

    private void ApplyFirstRowAuthoredWidths(IReadOnlyList<TableFormattingRow> rows, HtmlRenderBoxStyle tableStyle, double[] widths, double contentWidth) {
        if (rows.Count == 0) return;
        int column = 0;
        foreach (TableFormattingCell cell in rows[0].Cells) {
            int span = Math.Min(cell.ColumnSpan, widths.Length - column);
            if (span <= 0) break;
            HtmlRenderBoxStyle style = ResolveFormattingCellStyle(cell, contentWidth);
            if (style.ExplicitWidth.HasValue) {
                double authored = style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
                ApplySpanningWidth(widths, column, span, authored);
            }
            column += span;
            if (column >= widths.Length) break;
        }
    }

    private void ApplyDeclaredColumnWidths(IElement table, double contentWidth, HtmlRenderBoxStyle tableStyle, double[] minimums, double[] preferred, bool[]? authoredWidths = null) {
        int column = 0;
        foreach (IElement element in table.QuerySelectorAll("col").Where(candidate => BelongsToTableColumn(candidate, table))) {
            int span = Math.Min(ReadSpan(element.GetAttribute("span"), minimums.Length), minimums.Length - column);
            if (span <= 0) break;
            HtmlRenderBoxStyle style = _styleResolver.Resolve(element, contentWidth, tableStyle);
            if (style.ExplicitWidth.HasValue) {
                double width = Math.Max(1D, style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
                for (int offset = 0; offset < span; offset++) {
                    // A col span repeats this column's declaration; unlike a
                    // cell colspan, its width is not shared across the tracks.
                    minimums[column + offset] = Math.Max(minimums[column + offset], width);
                    preferred[column + offset] = Math.Max(preferred[column + offset], width);
                    if (authoredWidths != null && !style.ExplicitWidthUsesPercentage) authoredWidths[column + offset] = true;
                }
            }
            column += span;
            if (column >= minimums.Length) break;
        }
    }

    private void ResolveTableCellIntrinsicWidths(TableFormattingCell cell, HtmlRenderBoxStyle style, double containingWidth, int depth, out double minimum, out double preferred) {
        if (cell.Nodes == null && IsSpecializedTableCell(cell.Element)) {
            preferred = IsReplacedImageElement(cell.Element) || IsInputType(cell.Element, "image")
                ? ResolveIntrinsicReplacedImageBoxWidth(cell.Element, style)
                : ResolveFormControlIntrinsicOuterWidth(cell.Element, style, containingWidth);
            minimum = style.ExplicitWidthUsesPercentage || style.MaxWidthUsesPercentage
                ? ResolveCompressibleReplacedMinimumWidth(style) : preferred;
            return;
        }
        IReadOnlyList<IntrinsicTextRun> runs;
        if (cell.Nodes == null) {
            runs = ResolveInFlowIntrinsicTextRuns(new FlexItem(cell.Element, style, 0), containingWidth, depth,
                skipSizedNestedTables: true, includeDescendantInsets: true);
        } else {
            var rawRuns = new List<IntrinsicTextRun>();
            AppendInFlowIntrinsicTextRuns(cell.Element, style, containingWidth, depth, rawRuns,
                skipSizedNestedTables: true, includeDescendantInsets: true, sourceNodes: cell.Nodes);
            runs = NormalizeIntrinsicTextRuns(rawRuns);
        }
        double insets = style.HorizontalInsets;
        minimum = Math.Max(1D, MeasureMinContentRuns(runs) + insets);
        preferred = Math.Max(minimum, MeasureMaxContentRuns(runs) + insets);
        if (style.ExplicitWidth.HasValue) {
            double authored = style.ExplicitWidth.Value + (style.BorderBox ? 0D : insets);
            minimum = Math.Max(minimum, authored);
            preferred = Math.Max(preferred, authored);
        }
        if (style.MinWidth.HasValue) {
            double authoredMinimum = style.MinWidth.Value + (style.BorderBox ? 0D : insets);
            minimum = Math.Max(minimum, authoredMinimum);
            preferred = Math.Max(preferred, authoredMinimum);
        }
    }

    /// <summary>Applies spans after single-column contributions without source-row-order dependence.</summary>
    private static void ApplyTableSpanningWidths(double[] minimums, double[] preferred,
        IReadOnlyList<(int Column, int Span, double Minimum, double Preferred)> cells, double spacing) {
        foreach (var group in cells.GroupBy(cell => cell.Span).OrderBy(group => group.Key)) {
            // Equal-length spans all read the previous stage, so overlapping
            // cells cannot change the basis used by the next source row.
            double[] priorMinimums = minimums.ToArray();
            double[] priorPreferred = preferred.ToArray();
            foreach (var cell in group) {
                double internalSpacing = spacing * (cell.Span - 1);
                double minimumExtra = Math.Max(0D, cell.Minimum - internalSpacing
                    - SumColumnWidths(priorMinimums, cell.Column, cell.Span)) / cell.Span;
                double preferredExtra = Math.Max(0D, cell.Preferred - internalSpacing
                    - SumColumnWidths(priorPreferred, cell.Column, cell.Span)) / cell.Span;
                for (int offset = 0; offset < cell.Span; offset++) {
                    int column = cell.Column + offset;
                    minimums[column] = Math.Max(minimums[column], priorMinimums[column] + minimumExtra);
                    preferred[column] = Math.Max(preferred[column], priorPreferred[column] + preferredExtra);
                }
            }
            for (int column = 0; column < minimums.Length; column++) preferred[column] = Math.Max(preferred[column], minimums[column]);
        }
    }

    private static IReadOnlyList<double> AllocateFixedColumnWidths(IReadOnlyList<double> requested, double totalWidth) {
        var result = requested.Select(value => Math.Max(0D, value)).ToArray();
        int unspecified = result.Count(value => value <= 0.0001D);
        double specifiedTotal = result.Sum();
        if (specifiedTotal > totalWidth && specifiedTotal > 0D) {
            double scale = totalWidth / specifiedTotal;
            for (int index = 0; index < result.Length; index++) result[index] = result[index] > 0D ? result[index] * scale : 0.01D;
        } else if (unspecified > 0) {
            double share = Math.Max(0.01D, (totalWidth - specifiedTotal) / unspecified);
            for (int index = 0; index < result.Length; index++) if (result[index] <= 0.0001D) result[index] = share;
        } else if (result.Length > 0) {
            double extra = (totalWidth - specifiedTotal) / result.Length;
            for (int index = 0; index < result.Length; index++) result[index] += extra;
        }
        NormalizeColumnWidthTotal(result, totalWidth);
        return result;
    }

    private static IReadOnlyList<double> AllocateAutoColumnWidths(IReadOnlyList<double> minimums, IReadOnlyList<double> preferred, IReadOnlyList<bool> authoredWidths, double totalWidth) {
        var result = new double[minimums.Count];
        double minimumTotal = minimums.Sum();
        double preferredTotal = preferred.Sum();
        if (totalWidth <= minimumTotal + 0.0001D) {
            double scale = totalWidth / Math.Max(0.01D, minimumTotal);
            for (int index = 0; index < result.Length; index++) result[index] = Math.Max(0.01D, minimums[index] * scale);
        } else if (totalWidth < preferredTotal - 0.0001D) {
            double progress = (totalWidth - minimumTotal) / Math.Max(0.01D, preferredTotal - minimumTotal);
            for (int index = 0; index < result.Length; index++) result[index] = minimums[index] + (preferred[index] - minimums[index]) * progress;
        } else {
            int flexibleColumns = authoredWidths.Count(authored => !authored);
            bool distributeToAll = flexibleColumns == 0;
            double extra = (totalWidth - preferredTotal) / (distributeToAll ? result.Length : flexibleColumns);
            for (int index = 0; index < result.Length; index++) {
                result[index] = preferred[index] + (distributeToAll || !authoredWidths[index] ? extra : 0D);
            }
        }
        NormalizeColumnWidthTotal(result, totalWidth);
        return result;
    }

    private static void ApplySpanningWidth(double[] widths, int start, int span, double required) {
        double current = SumColumnWidths(widths, start, span);
        double deficit = required - current;
        if (deficit <= 0.0001D) return;
        double addition = deficit / span;
        for (int offset = 0; offset < span; offset++) widths[start + offset] += addition;
    }

    private static void NormalizeColumnWidthTotal(double[] widths, double totalWidth) {
        if (widths.Length == 0) return;
        widths[widths.Length - 1] += totalWidth - widths.Sum();
        widths[widths.Length - 1] = Math.Max(0.01D, widths[widths.Length - 1]);
    }

    private static double[] CreateColumnOffsets(IReadOnlyList<double> widths) {
        var offsets = new double[widths.Count];
        for (int index = 1; index < offsets.Length; index++) offsets[index] = offsets[index - 1] + widths[index - 1];
        return offsets;
    }

    private static double SumColumnWidths(IReadOnlyList<double> widths, int start, int count) {
        double result = 0D;
        for (int index = start; index < start + count && index < widths.Count; index++) result += widths[index];
        return result;
    }

    private int DetermineDeclaredColumnCount(IElement table) {
        long count = 0;
        foreach (IElement element in table.QuerySelectorAll("col").Where(candidate => BelongsToTableColumn(candidate, table))) {
            count += ReadSpan(element.GetAttribute("span"), 1000);
            EnsureTableColumnLimit(count);
        }
        return (int)count;
    }

    private static bool BelongsToTableColumn(IElement column, IElement table) {
        IElement? current = column.ParentElement;
        while (current != null && !string.Equals(current.TagName, "table", StringComparison.OrdinalIgnoreCase)) current = current.ParentElement;
        return ReferenceEquals(current, table);
    }
}
