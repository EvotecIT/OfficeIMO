using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<double> ResolveTableColumnWidths(IReadOnlyList<TableFormattingRow> rows, IElement table, int columnCount, double contentWidth, double captionMinimumWidth, HtmlRenderBoxStyle tableStyle, int depth, int legacyBorderWidth, out double usedWidth) {
        if (tableStyle.TableLayout == "fixed") {
            return ResolveFixedTableColumnWidths(rows, table, columnCount, contentWidth, tableStyle, out usedWidth);
        }

        ResolveAutoTableColumnContributions(rows, table, columnCount, contentWidth, tableStyle, depth, legacyBorderWidth,
            out double[] minimums, out double[] preferred, out bool[] authoredWidths, out double[] percentages,
            out double[] contentMinimums);
        double minimumAllowed = tableStyle.MinWidth.HasValue
            ? Math.Max(0.01D, tableStyle.MinWidth.Value - (tableStyle.BorderBox ? tableStyle.HorizontalInsets : 0D))
            : 0.01D;
        minimumAllowed = Math.Max(minimumAllowed, captionMinimumWidth);
        usedWidth = tableStyle.ExplicitWidth.HasValue
            ? contentWidth
            : Math.Min(contentWidth, Math.Max(minimumAllowed, preferred.Sum()));
        // Percentage requests compete within the used width; they are not
        // unbreakable content and must not force an over-constrained table wider.
        usedWidth = Math.Max(usedWidth, (percentages.Any(value => value > 0D) ? contentMinimums : minimums).Sum());
        return AllocateAutoColumnWidths(minimums, preferred, authoredWidths, percentages, contentMinimums, usedWidth);
    }

    /// <summary>Collects column constraints once for both table layout and non-paint intrinsic sizing.</summary>
    private void ResolveAutoTableColumnContributions(IReadOnlyList<TableFormattingRow> rows, IElement table,
        int columnCount, double contentWidth, HtmlRenderBoxStyle tableStyle, int depth, int legacyBorderWidth,
        out double[] minimums, out double[] preferred, out bool[] authoredWidths, out double[] percentages,
        out double[] contentMinimums) {
        minimums = new double[columnCount];
        preferred = new double[columnCount];
        authoredWidths = new bool[columnCount];
        percentages = new double[columnCount];
        contentMinimums = new double[columnCount];
        ApplyDeclaredColumnWidths(table, contentWidth, tableStyle, minimums, preferred, authoredWidths, percentageRequests: percentages);
        var contentPreferred = new double[columnCount];
        for (int index = 0; index < columnCount; index++) {
            if (percentages[index] <= 0D) contentPreferred[index] = preferred[index];
        }
        var spanningCells = new List<(int Column, int Span, double Minimum, double Preferred)>();
        var spanningContent = new List<(int Column, int Span, double Minimum, double Preferred)>();
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
                if (span == 1 && cellStyle.ExplicitWidth.HasValue) {
                    if (cellStyle.ExplicitWidthUsesPercentage) {
                        percentages[column] = Math.Max(percentages[column], cellStyle.ExplicitWidth.Value / Math.Max(0.01D, contentWidth));
                    } else authoredWidths[column] = true;
                }
                ResolveTableCellIntrinsicWidths(cell, cellStyle, contentWidth, depth,
                    out double minimum, out double maximum, out double contentMinimum, out double contentMaximum);
                if (span == 1) {
                    minimums[column] = Math.Max(minimums[column], minimum);
                    preferred[column] = Math.Max(preferred[column], maximum);
                    contentMinimums[column] = Math.Max(contentMinimums[column], contentMinimum);
                    contentPreferred[column] = Math.Max(contentPreferred[column], contentMaximum);
                } else {
                    spanningCells.Add((column, span, minimum, maximum));
                    spanningContent.Add((column, span,
                        cellStyle.ExplicitWidthUsesPercentage ? contentMinimum : minimum, contentMaximum));
                }
                int rowSpan = ReadRowSpan(cell.RowSpan, rows, rowIndex);
                for (int occupied = column; occupied < column + span; occupied++) occupancy[occupied] = Math.Max(occupancy[occupied], rowSpan);
                column += span;
            }
            DecrementOccupancy(occupancy);
        }
        double spacing = tableStyle.BorderCollapse == "collapse" ? 0D : tableStyle.BorderSpacingX;
        ApplyTableSpanningContentWidths(contentMinimums, contentPreferred, spanningContent, spacing, authoredWidths, percentages);
        for (int index = 0; index < columnCount; index++) {
            minimums[index] = Math.Max(minimums[index], contentMinimums[index]);
            preferred[index] = Math.Max(preferred[index], contentPreferred[index]);
        }
        ApplyTableSpanningWidths(minimums, preferred, spanningCells, spacing);
    }

    /// <summary>Measures a table as one atomic box using source columns, without painting or feeding layout widths back into measurement.</summary>
    private (double Minimum, double Maximum) ResolveIntrinsicTableOuterWidths(IElement table,
        HtmlRenderBoxStyle style, double availableWidth, int depth) {
        EnsureDepth(depth, table);
        style = _styleResolver.ResolveIntrinsicMeasurementStyle(table, style);
        if (style.TableLayout == "fixed" && style.ExplicitWidth.HasValue && !style.ExplicitWidthUsesPercentage) {
            double fixedWidth = ResolveFixedTableIntrinsicOuterWidth(table, style, availableWidth, depth);
            return (fixedWidth, fixedWidth);
        }
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
                out double[] minimums, out double[] preferred, out _, out double[] percentages, out double[] contentMinimums);
            double spacing = style.BorderCollapse == "collapse" ? 0D : style.BorderSpacingX * (columnCount + 1);
            minimum = Math.Max(minimum, (percentages.Any(value => value > 0D) ? contentMinimums : minimums).Sum() + spacing);
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

    private void ApplyDeclaredColumnWidths(IElement table, double contentWidth, HtmlRenderBoxStyle tableStyle, double[] minimums, double[] preferred, bool[]? authoredWidths = null, bool[]? declaredWidths = null, bool[]? percentageWidths = null, double[]? percentageRequests = null) {
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
                    if (declaredWidths != null) declaredWidths[column + offset] = true;
                    if (percentageWidths != null) percentageWidths[column + offset] = style.ExplicitWidthUsesPercentage;
                    if (percentageRequests != null && style.ExplicitWidthUsesPercentage) {
                        percentageRequests[column + offset] = Math.Max(percentageRequests[column + offset], style.ExplicitWidth.Value / Math.Max(0.01D, contentWidth));
                    }
                }
            }
            column += span;
            if (column >= minimums.Length) break;
        }
    }

    private void ResolveTableCellIntrinsicWidths(TableFormattingCell cell, HtmlRenderBoxStyle style, double containingWidth,
        int depth, out double minimum, out double preferred, out double contentMinimum, out double contentPreferred) {
        if (cell.Nodes == null && IsSpecializedTableCell(cell.Element)) {
            preferred = IsReplacedImageElement(cell.Element) || IsInputType(cell.Element, "image")
                ? ResolveIntrinsicReplacedImageBoxWidth(cell.Element, style)
                : ResolveFormControlIntrinsicOuterWidth(cell.Element, style, containingWidth);
            minimum = style.ExplicitWidthUsesPercentage || style.MaxWidthUsesPercentage
                ? ResolveCompressibleReplacedMinimumWidth(style) : preferred;
            contentMinimum = minimum;
            contentPreferred = preferred;
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
        // Empty content contributes only its authored edges, not a synthetic pixel.
        minimum = Math.Max(0D, (runs.Count == 0 ? 0D : MeasureMinContentRuns(runs)) + insets);
        contentMinimum = minimum;
        preferred = Math.Max(minimum, (runs.Count == 0 ? 0D : MeasureMaxContentRuns(runs)) + insets);
        contentPreferred = preferred;
        if (style.ExplicitWidth.HasValue) {
            double authored = style.ExplicitWidth.Value + (style.BorderBox || style.ExplicitWidthUsesPercentage ? 0D : insets);
            minimum = Math.Max(minimum, authored);
            preferred = Math.Max(preferred, authored);
            if (!style.ExplicitWidthUsesPercentage) contentPreferred = Math.Max(contentPreferred, authored);
        }
        if (style.MinWidth.HasValue) {
            double authoredMinimum = style.MinWidth.Value + (style.BorderBox ? 0D : insets);
            minimum = Math.Max(minimum, authoredMinimum);
            preferred = Math.Max(preferred, authoredMinimum);
            contentMinimum = Math.Max(contentMinimum, authoredMinimum);
            contentPreferred = Math.Max(contentPreferred, authoredMinimum);
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

    private static IReadOnlyList<double> AllocateAutoColumnWidths(IReadOnlyList<double> minimums, IReadOnlyList<double> preferred,
        IReadOnlyList<bool> authoredWidths, IReadOnlyList<double> percentages, IReadOnlyList<double> contentMinimums, double totalWidth) {
        var result = new double[minimums.Count];
        double minimumTotal = minimums.Sum();
        double preferredTotal = preferred.Sum();
        if (minimumTotal > totalWidth + 0.0001D && percentages.Any(value => value > 0D)) {
            return AllocateOverconstrainedAutoColumns(contentMinimums, preferred, percentages, totalWidth);
        }
        if (totalWidth <= minimumTotal + 0.0001D) {
            double scale = totalWidth / Math.Max(0.01D, minimumTotal);
            for (int index = 0; index < result.Length; index++) result[index] = Math.Max(0D, minimums[index] * scale);
        } else if (totalWidth < preferredTotal - 0.0001D) {
            double progress = (totalWidth - minimumTotal) / Math.Max(0.01D, preferredTotal - minimumTotal);
            for (int index = 0; index < result.Length; index++) result[index] = minimums[index] + (preferred[index] - minimums[index]) * progress;
        } else {
            // Content-sized columns take surplus first. Explicit lengths follow,
            // then percentages; scaling every column would grow a reserved share.
            bool hasContentColumns = Enumerable.Range(0, result.Length).Any(index => !authoredWidths[index] && percentages[index] <= 0D);
            bool hasNonPercentageColumns = Enumerable.Range(0, result.Length).Any(index => percentages[index] <= 0D);
            bool Eligible(int index) => hasContentColumns
                ? !authoredWidths[index] && percentages[index] <= 0D
                : !hasNonPercentageColumns || percentages[index] <= 0D;
            double Weight(int index) => hasNonPercentageColumns ? preferred[index] : percentages[index];
            double weightTotal = Enumerable.Range(0, result.Length).Where(Eligible).Sum(Weight);
            int count = Enumerable.Range(0, result.Length).Count(Eligible);
            double extra = totalWidth - preferredTotal;
            for (int index = 0; index < result.Length; index++) {
                double share = !Eligible(index) ? 0D : weightTotal > 0D ? Weight(index) / weightTotal : 1D / count;
                result[index] = preferred[index] + extra * share;
            }
        }
        NormalizeColumnWidthTotal(result, totalWidth, allowZeroWidth: true);
        return result;
    }

    /// <summary>Fits competing percentage hints in source-column order while retaining every content minimum.</summary>
    private static IReadOnlyList<double> AllocateOverconstrainedAutoColumns(IReadOnlyList<double> minimums,
        IReadOnlyList<double> preferred, IReadOnlyList<double> percentages, double totalWidth) {
        double[] result = minimums.ToArray();
        double remaining = Math.Max(0D, totalWidth - result.Sum());
        for (int index = 0; index < result.Length; index++) {
            if (percentages[index] <= 0D) continue;
            double extra = Math.Min(remaining, Math.Max(0D, percentages[index] * totalWidth - result[index]));
            result[index] += extra;
            remaining -= extra;
        }
        int[] nonPercentage = Enumerable.Range(0, result.Length).Where(index => percentages[index] <= 0D).ToArray();
        double requested = nonPercentage.Sum(index => Math.Max(0D, preferred[index] - result[index]));
        double available = Math.Min(remaining, requested);
        if (requested > 0D) {
            foreach (int index in nonPercentage) result[index] += available * Math.Max(0D, preferred[index] - result[index]) / requested;
            remaining -= available;
        }
        if (remaining > 0D) {
            int[] recipients = nonPercentage.Length > 0 ? nonPercentage : Enumerable.Range(0, result.Length).ToArray();
            double weight = recipients.Sum(index => preferred[index]);
            foreach (int index in recipients) result[index] += remaining * (weight > 0D ? preferred[index] / weight : 1D / recipients.Length);
        }
        NormalizeColumnWidthTotal(result, totalWidth, allowZeroWidth: true);
        return result;
    }

    private static void NormalizeColumnWidthTotal(double[] widths, double totalWidth, bool allowZeroWidth = false) {
        if (widths.Length == 0) return;
        widths[widths.Length - 1] += totalWidth - widths.Sum();
        widths[widths.Length - 1] = Math.Max(allowZeroWidth ? 0D : 0.01D, widths[widths.Length - 1]);
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
