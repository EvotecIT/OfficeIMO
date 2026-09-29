using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<double> ResolveTableColumnWidths(IReadOnlyList<IElement> rows, IReadOnlyDictionary<IElement, HtmlRenderBoxStyle> rowStyles, IElement table, int columnCount, double contentWidth, double captionMinimumWidth, HtmlRenderBoxStyle tableStyle, int depth, out double usedWidth) {
        if (tableStyle.TableLayout == "fixed") {
            var fixedWidths = new double[columnCount];
            ApplyDeclaredColumnWidths(table, contentWidth, tableStyle, fixedWidths, fixedWidths);
            ApplyFirstRowAuthoredWidths(rows, tableStyle, fixedWidths, contentWidth);
            usedWidth = contentWidth;
            return AllocateFixedColumnWidths(fixedWidths, contentWidth);
        }

        var minimums = Enumerable.Repeat(1D, columnCount).ToArray();
        var preferred = Enumerable.Repeat(1D, columnCount).ToArray();
        var authoredWidths = new bool[columnCount];
        ApplyDeclaredColumnWidths(table, contentWidth, tableStyle, minimums, preferred, authoredWidths);
        var occupancy = new int[columnCount];
        int legacyBorderWidth = ReadLegacyTableBorderWidth(table);
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            int column = 0;
            foreach (IElement cell in rows[rowIndex].Children.Where(IsTableCell)) {
                int requestedSpan = ReadSpan(cell.GetAttribute("colspan"), columnCount);
                column = FindAvailableColumn(occupancy, column, requestedSpan);
                if (column >= columnCount) break;
                int span = Math.Max(1, Math.Min(requestedSpan, columnCount - column));
                HtmlRenderBoxStyle cellStyle = _styleResolver.Resolve(cell, contentWidth, rowStyles[rows[rowIndex]]);
                ApplyTableCellFallbackInsets(cellStyle, legacyBorderWidth);
                if (span == 1 && cellStyle.ExplicitWidth.HasValue && !cellStyle.ExplicitWidthUsesPercentage) {
                    authoredWidths[column] = true;
                }
                ResolveTableCellIntrinsicWidths(cell, cellStyle, contentWidth, depth, out double minimum, out double maximum);
                ApplySpanningWidth(minimums, column, span, minimum);
                ApplySpanningWidth(preferred, column, span, maximum);
                int rowSpan = ReadRowSpan(cell.GetAttribute("rowspan"), rows, rowIndex, table);
                for (int occupied = column; occupied < column + span; occupied++) occupancy[occupied] = Math.Max(occupancy[occupied], rowSpan);
                column += span;
            }
            DecrementOccupancy(occupancy);
        }
        double minimumAllowed = tableStyle.MinWidth.HasValue
            ? Math.Max(0.01D, tableStyle.MinWidth.Value - (tableStyle.BorderBox ? tableStyle.HorizontalInsets : 0D))
            : 0.01D;
        minimumAllowed = Math.Max(minimumAllowed, captionMinimumWidth);
        usedWidth = tableStyle.ExplicitWidth.HasValue
            ? contentWidth
            : Math.Min(contentWidth, Math.Max(minimumAllowed, preferred.Sum()));
        return AllocateAutoColumnWidths(minimums, preferred, authoredWidths, usedWidth);
    }

    private double MeasureTableCaptionMinimumWidth(IElement table, double containingWidth, HtmlRenderBoxStyle tableStyle) {
        IElement? caption = table.Children.FirstOrDefault(child => string.Equals(child.TagName, "caption", StringComparison.OrdinalIgnoreCase));
        if (caption == null) return 0D;
        HtmlRenderBoxStyle captionStyle = _styleResolver.Resolve(caption, containingWidth, tableStyle);
        if (captionStyle.Display == "none") return 0D;
        IReadOnlyList<GridIntrinsicTextRun> runs = ResolveGridInFlowTextRuns(new FlexItem(caption, captionStyle, 0), containingWidth);
        double minimum = MeasureGridMinContentRuns(runs) + captionStyle.HorizontalInsets;
        if (captionStyle.ExplicitWidth.HasValue && !captionStyle.ExplicitWidthUsesPercentage) {
            minimum = Math.Max(minimum, captionStyle.ExplicitWidth.Value
                + (captionStyle.BorderBox ? 0D : captionStyle.HorizontalInsets));
        }
        return minimum + captionStyle.MarginLeft + captionStyle.MarginRight;
    }

    private void ApplyFirstRowAuthoredWidths(IReadOnlyList<IElement> rows, HtmlRenderBoxStyle tableStyle, double[] widths, double contentWidth) {
        if (rows.Count == 0) return;
        int column = 0;
        foreach (IElement cell in rows[0].Children.Where(IsTableCell)) {
            int span = Math.Min(ReadSpan(cell.GetAttribute("colspan"), widths.Length), widths.Length - column);
            if (span <= 0) break;
            HtmlRenderBoxStyle style = _styleResolver.Resolve(cell, contentWidth, tableStyle);
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
                double perColumn = width / span;
                for (int offset = 0; offset < span; offset++) {
                    minimums[column + offset] = Math.Max(minimums[column + offset], perColumn);
                    preferred[column + offset] = Math.Max(preferred[column + offset], perColumn);
                    if (authoredWidths != null && !style.ExplicitWidthUsesPercentage) authoredWidths[column + offset] = true;
                }
            }
            column += span;
            if (column >= minimums.Length) break;
        }
    }

    private void ResolveTableCellIntrinsicWidths(IElement cell, HtmlRenderBoxStyle style, double containingWidth, int depth, out double minimum, out double preferred) {
        string text = ApplyTextTransform(cell.TextContent ?? string.Empty, style);
        // Cell content is text, not a CSS component value: quotes and parentheses
        // must not protect whitespace (including tabs) from intrinsic sizing.
        IReadOnlyList<string> tokens = text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        string normalized = string.Join(" ", tokens);
        double insets = style.HorizontalInsets;
        minimum = tokens.Count == 0 ? insets + 1D : tokens.Max(token => MeasureInlineText(token, style)) + insets;
        preferred = Math.Max(minimum, MeasureInlineText(normalized, style) + insets);
        bool hasLineBreak = ContainsTableCellLineBreak(cell);
        if (hasLineBreak) {
            // TextContent drops forced line breaks. Use the existing in-flow run
            // measurement so a navigation cell with <br> does not reserve one
            // column as though all of its labels occupied a single line.
            IReadOnlyList<GridIntrinsicTextRun> runs = ResolveGridInFlowTextRuns(new FlexItem(cell, style, 0), containingWidth);
            minimum = Math.Max(1D, MeasureGridMinContentRuns(runs) + insets);
            preferred = Math.Max(minimum, MeasureGridMaxContentRuns(runs) + insets);
        }
        if (!hasLineBreak && text.IndexOf('\t') >= 0) {
            preferred = Math.Max(preferred, MeasureTabExpandedText(text, style, 0D) + insets);
        }
        if (!hasLineBreak) {
            // The in-flow runs already include generated content when a cell
            // contains a forced break; do not count it a second time.
            double generatedPreferred = 0D;
            double generatedMinimum = 0D;
            MeasureTableCellGeneratedContent(cell, HtmlPseudoElementKind.Before, style, containingWidth, ref generatedMinimum, ref generatedPreferred);
            MeasureTableCellGeneratedContent(cell, HtmlPseudoElementKind.After, style, containingWidth, ref generatedMinimum, ref generatedPreferred);
            minimum = Math.Max(minimum, generatedMinimum + insets);
            preferred += generatedPreferred;
        }
        if (style.ExplicitWidth.HasValue) {
            double authored = style.ExplicitWidth.Value + (style.BorderBox ? 0D : insets);
            minimum = Math.Max(minimum, authored);
            preferred = Math.Max(preferred, authored);
        }
        ResolveTableDescendantIntrinsicWidths(cell, style, containingWidth, depth, insets, ref minimum, ref preferred);
    }

    private static bool ContainsTableCellLineBreak(IElement cell) {
        if (cell.Children.Length == 0) return false;
        var pending = new Stack<IElement>();
        foreach (IElement child in cell.Children) pending.Push(child);
        while (pending.Count > 0) {
            IElement element = pending.Pop();
            if (string.Equals(element.LocalName, "br", StringComparison.OrdinalIgnoreCase)) return true;
            if (string.Equals(element.LocalName, "table", StringComparison.OrdinalIgnoreCase)) continue;
            foreach (IElement child in element.Children) pending.Push(child);
        }
        return false;
    }

    private void MeasureTableCellGeneratedContent(IElement cell, HtmlPseudoElementKind kind, HtmlRenderBoxStyle cellStyle,
        double containingWidth, ref double minimum, ref double preferred) {
        var runs = new List<HtmlInlineRun>();
        AddGeneratedInlineRun(cell, kind, containingWidth, null, cellStyle, null, 0D, 0D, runs);
        foreach (HtmlInlineRun run in runs) {
            if (run.AtomicBlock != null) {
                minimum = Math.Max(minimum, run.AtomicBlock.Width);
                preferred += run.AtomicBlock.Width;
                continue;
            }
            if (run.Text.Length == 0) continue;
            string[] tokens = run.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length > 0) minimum = Math.Max(minimum, tokens.Max(token => MeasureInlineText(token, run.Style)));
            preferred += MeasureInlineText(run.Text, run.Style);
        }
    }

    private void ResolveTableDescendantIntrinsicWidths(
        IElement parent,
        HtmlRenderBoxStyle parentStyle,
        double containingWidth,
        int depth,
        double cellInsets,
        ref double minimum,
        ref double preferred) {
        foreach (IElement element in parent.Children) {
            CheckCancellation();
            EnsureDepth(depth + 1, element);
            ChargeLayoutOperation(HtmlRenderStyleResolver.DescribeSource(element));
            HtmlRenderBoxStyle elementStyle = _styleResolver.Resolve(element, containingWidth, parentStyle);
            if (string.Equals(elementStyle.Display, "none", StringComparison.OrdinalIgnoreCase)) continue;
            if (ShouldExtractOutOfFlow(elementStyle)) continue;

            double descendantWidth = 0D;
            if (string.Equals(element.TagName, "img", StringComparison.OrdinalIgnoreCase)
                || string.Equals(element.TagName, "svg", StringComparison.OrdinalIgnoreCase)) {
                if (elementStyle.ExplicitWidthUsesPercentage) continue;
                descendantWidth = ResolveReplacedImageBoxWidth(element, elementStyle);
            } else if (elementStyle.ExplicitWidth.HasValue && !elementStyle.ExplicitWidthUsesPercentage) {
                descendantWidth = elementStyle.ExplicitWidth.Value
                    + (elementStyle.BorderBox ? 0D : elementStyle.HorizontalInsets);
            }
            if (descendantWidth > 0D) {
                descendantWidth += elementStyle.MarginLeft + elementStyle.MarginRight + cellInsets;
                minimum = Math.Max(minimum, descendantWidth);
                preferred = Math.Max(preferred, descendantWidth);
            }

            if (!IsTableCell(element)) {
                ResolveTableDescendantIntrinsicWidths(
                    element,
                    elementStyle,
                    containingWidth,
                    depth + 1,
                    cellInsets,
                    ref minimum,
                    ref preferred);
            }
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
