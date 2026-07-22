using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<double> ResolveTableColumnWidths(IReadOnlyList<IElement> rows, IElement table, int columnCount, double contentWidth, HtmlRenderBoxStyle tableStyle) {
        if (tableStyle.TableLayout == "fixed") {
            var fixedWidths = new double[columnCount];
            ApplyDeclaredColumnWidths(table, contentWidth, tableStyle, fixedWidths, fixedWidths);
            ApplyFirstRowAuthoredWidths(rows, tableStyle, fixedWidths, contentWidth);
            return AllocateFixedColumnWidths(fixedWidths, contentWidth);
        }

        var minimums = Enumerable.Repeat(1D, columnCount).ToArray();
        var preferred = Enumerable.Repeat(1D, columnCount).ToArray();
        ApplyDeclaredColumnWidths(table, contentWidth, tableStyle, minimums, preferred);
        var occupancy = new int[columnCount];
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            int column = 0;
            foreach (IElement cell in rows[rowIndex].Children.Where(IsTableCell)) {
                int requestedSpan = ReadSpan(cell.GetAttribute("colspan"), columnCount);
                column = FindAvailableColumn(occupancy, column, requestedSpan);
                if (column >= columnCount) break;
                int span = Math.Max(1, Math.Min(requestedSpan, columnCount - column));
                HtmlRenderBoxStyle cellStyle = _styleResolver.Resolve(cell, contentWidth, tableStyle);
                ResolveTableCellIntrinsicWidths(cell, cellStyle, contentWidth, out double minimum, out double maximum);
                ApplySpanningWidth(minimums, column, span, minimum);
                ApplySpanningWidth(preferred, column, span, maximum);
                int rowSpan = ReadRowSpan(cell.GetAttribute("rowspan"), rows, rowIndex, table);
                for (int occupied = column; occupied < column + span; occupied++) occupancy[occupied] = Math.Max(occupancy[occupied], rowSpan);
                column += span;
            }
            DecrementOccupancy(occupancy);
        }
        return AllocateAutoColumnWidths(minimums, preferred, contentWidth);
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

    private void ApplyDeclaredColumnWidths(IElement table, double contentWidth, HtmlRenderBoxStyle tableStyle, double[] minimums, double[] preferred) {
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
                }
            }
            column += span;
            if (column >= minimums.Length) break;
        }
    }

    private void ResolveTableCellIntrinsicWidths(IElement cell, HtmlRenderBoxStyle style, double containingWidth, out double minimum, out double preferred) {
        string text = ApplyTextTransform(cell.TextContent ?? string.Empty, style.TextTransform);
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(text);
        string normalized = string.Join(" ", tokens);
        double insets = style.HorizontalInsets;
        minimum = tokens.Count == 0 ? insets + 1D : tokens.Max(token => MeasureText(token, style.Font)) + insets;
        preferred = Math.Max(minimum, MeasureText(normalized, style.Font) + insets);
        if (style.ExplicitWidth.HasValue) {
            double authored = style.ExplicitWidth.Value + (style.BorderBox ? 0D : insets);
            minimum = Math.Max(minimum, authored);
            preferred = Math.Max(preferred, authored);
        }
        foreach (IElement image in cell.QuerySelectorAll("img").Where(candidate => BelongsToTableCell(candidate, cell))) {
            HtmlRenderBoxStyle imageStyle = _styleResolver.Resolve(image, containingWidth, style);
            double imageWidth = ResolveReplacedImageBoxWidth(image, imageStyle) + imageStyle.MarginLeft + imageStyle.MarginRight + insets;
            minimum = Math.Max(minimum, imageWidth);
            preferred = Math.Max(preferred, imageWidth);
        }
    }

    private static bool BelongsToTableCell(IElement element, IElement cell) {
        IElement? current = element.ParentElement;
        while (current != null && !IsTableCell(current)) current = current.ParentElement;
        return ReferenceEquals(current, cell);
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

    private static IReadOnlyList<double> AllocateAutoColumnWidths(IReadOnlyList<double> minimums, IReadOnlyList<double> preferred, double totalWidth) {
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
            double extra = (totalWidth - preferredTotal) / result.Length;
            for (int index = 0; index < result.Length; index++) result[index] = preferred[index] + extra;
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
