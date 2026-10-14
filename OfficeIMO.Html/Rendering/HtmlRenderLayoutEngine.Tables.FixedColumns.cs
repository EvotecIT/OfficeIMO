using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Resolves fixed tracks once; absolute declarations enlarge the used table instead of being compressed.</summary>
    private IReadOnlyList<double> ResolveFixedTableColumnWidths(IReadOnlyList<TableFormattingRow> rows, IElement table,
        int columnCount, double contentWidth, HtmlRenderBoxStyle tableStyle, out double usedWidth) {
        var widths = new double[columnCount];
        var declaredColumns = new bool[columnCount];
        var percentageWidths = new bool[columnCount];
        ApplyDeclaredColumnWidths(table, contentWidth, tableStyle, widths, widths,
            declaredWidths: declaredColumns, percentageWidths: percentageWidths);
        ApplyFirstRowAuthoredWidths(rows, widths, declaredColumns, percentageWidths, contentWidth,
            tableStyle.BorderCollapse == "collapse" ? 0D : tableStyle.BorderSpacingX);
        double absoluteMinimum = 0D;
        int flexibleTracks = 0;
        for (int column = 0; column < widths.Length; column++) {
            if (percentageWidths[column] || widths[column] <= 0.0001D) flexibleTracks++;
            else absoluteMinimum += widths[column];
        }
        // Percentage requests use the original table basis. Preserve absolute
        // tracks first, then fit percentages into the remaining used width.
        usedWidth = Math.Max(contentWidth, absoluteMinimum + flexibleTracks * 0.01D);
        return AllocateFixedColumnWidths(widths, percentageWidths, usedWidth);
    }

    /// <summary>Uses the same fixed-track constraints for nested atomic measurement and subsequent paint.</summary>
    private double ResolveFixedTableIntrinsicOuterWidth(IElement table, HtmlRenderBoxStyle style, double availableWidth, int depth) {
        style = style.Clone();
        PrepareTableBoxStyleForSizing(table, style, anonymous: false);
        double tableWidth = ResolveBoxWidth(Math.Max(1D, availableWidth - style.MarginLeft - style.MarginRight), style);
        double contentWidth = Math.Max(1D, tableWidth - style.HorizontalInsets);
        TableFormattingStructure formatting = BuildTableFormattingStructure(table, contentWidth, style, depth);
        int rowColumnCount = DetermineColumnCount(formatting.SizingRows);
        // Painting keeps an empty table at its prepared authored width even if
        // it declares columns; its intrinsic contribution must use that width.
        if (rowColumnCount > 0) {
            int columnCount = Math.Max(rowColumnCount, DetermineDeclaredColumnCount(table));
            double spacing = style.BorderCollapse == "collapse" ? 0D : style.BorderSpacingX * (columnCount + 1);
            ResolveFixedTableColumnWidths(formatting.SizingRows, table, columnCount, Math.Max(0.01D, contentWidth - spacing), style,
                out double usedTrackWidth);
            tableWidth = Math.Max(tableWidth, usedTrackWidth + spacing + style.HorizontalInsets);
        }
        return tableWidth + style.MarginLeft + style.MarginRight;
    }

    /// <summary>First-row hints fill only tracks without a column declaration; internal spacing is counted once.</summary>
    private void ApplyFirstRowAuthoredWidths(IReadOnlyList<TableFormattingRow> rows, double[] widths,
        IReadOnlyList<bool> declaredColumns, bool[] percentageWidths, double contentWidth, double spacing) {
        if (rows.Count == 0) return;
        int column = 0;
        foreach (TableFormattingCell cell in rows[0].Cells) {
            int span = Math.Min(cell.ColumnSpan, widths.Length - column);
            if (span <= 0) break;
            HtmlRenderBoxStyle style = ResolveFormattingCellStyle(cell, contentWidth);
            if (style.ExplicitWidth.HasValue) {
                double authored = style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
                double share = Math.Max(0D, authored - spacing * (span - 1)) / span;
                for (int offset = 0; offset < span; offset++) {
                    int track = column + offset;
                    if (declaredColumns[track]) continue;
                    widths[track] = share;
                    percentageWidths[track] = style.ExplicitWidthUsesPercentage;
                }
            }
            column += span;
            if (column >= widths.Length) break;
        }
    }

    private static IReadOnlyList<double> AllocateFixedColumnWidths(IReadOnlyList<double> requested,
        IReadOnlyList<bool> percentageWidths, double totalWidth) {
        var result = requested.Select(value => Math.Max(0D, value)).ToArray();
        int unspecified = result.Count(value => value <= 0.0001D);
        double specifiedTotal = result.Sum();
        if (specifiedTotal > totalWidth && specifiedTotal > 0D) {
            double absoluteTotal = 0D;
            var percentages = new List<int>();
            for (int index = 0; index < result.Length; index++) {
                if (result[index] <= 0.0001D) result[index] = 0.01D;
                else if (percentageWidths[index]) percentages.Add(index);
                else absoluteTotal += result[index];
            }
            double available = Math.Max(0D, totalWidth - absoluteTotal - unspecified * 0.01D);
            while (percentages.Count > 0) {
                double requestedPercentageTotal = percentages.Sum(index => requested[index]);
                double scale = available / requestedPercentageTotal;
                int floor = percentages.FindIndex(index => requested[index] * scale < 0.01D);
                if (floor < 0) {
                    foreach (int index in percentages) result[index] = requested[index] * scale;
                    break;
                }
                result[percentages[floor]] = 0.01D;
                percentages.RemoveAt(floor);
                available = Math.Max(0D, available - 0.01D);
            }
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

    /// <summary>Centers the final used border box after column constraints, including continuation reflow.</summary>
    private void ResolveUsedTableHorizontalAutoMargins(IElement table, HtmlRenderBoxStyle style, double containingWidth,
        double tableWidth, bool anonymous) {
        if (!style.MarginLeftAuto && !style.MarginRightAuto) return;
        double freeSpace = Math.Max(0D, containingWidth - style.MarginLeft - style.MarginRight - tableWidth);
        ApplyHorizontalAutoMarginFreeSpace(style, freeSpace);
        if (!anonymous) _layoutStyles[table] = style.Clone();
    }
}
