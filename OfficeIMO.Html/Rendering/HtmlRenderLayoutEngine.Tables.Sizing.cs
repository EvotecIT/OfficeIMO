namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Unlike a normal block, a table or cell's authored height is a minimum.
    private static double ResolveTableMinimumHeight(HtmlRenderBoxStyle style) =>
        Math.Max(Math.Max(1D, style.VerticalInsets), (style.ExplicitHeight ?? 0D) + (style.BorderBox ? 0D : style.VerticalInsets));

    private static double ResolveTableCellMinimumHeight(HtmlRenderBoxStyle style, HtmlInlineLayout inline) =>
        Math.Max(Math.Max(style.LineHeight, inline.Height) + style.VerticalInsets,
            (style.ExplicitHeight ?? 0D) + (style.BorderBox ? 0D : style.VerticalInsets));

    private void ApplyTableMinimumHeight(IReadOnlyList<TableRowLayout> rows, HtmlRenderBoxStyle style,
        double spacing) {
        if (rows.Count == 0 || !style.ExplicitHeight.HasValue) return;
        if (ApplyTablePercentageHeights(rows, style, spacing)) return;
        double naturalHeight = style.VerticalInsets + rows.Sum(row => row.Height) + spacing * (rows.Count + 1);
        double extra = Math.Max(0D, ResolveTableMinimumHeight(style) - naturalHeight) / rows.Count;
        foreach (TableRowLayout row in rows) row.Height += extra;
    }

    private static void ResolveTableBaselines(IReadOnlyList<TableRowLayout> rows) {
        foreach (TableRowLayout row in rows) {
            double descent = 0D;
            foreach (TableCellLayout cell in row.Cells) {
                if (!IsBaselineTableCell(cell)) continue;
                double baseline = ResolveTableCellBaseline(cell);
                row.Baseline = Math.Max(row.Baseline, baseline);
                if (cell.RowSpan == 1) descent = Math.Max(descent, cell.MinimumHeight - baseline);
            }
            row.Height = Math.Max(row.Height, row.Baseline + descent);
        }
    }

    private static void ResolveTableCellContentOffsets(IReadOnlyList<TableRowLayout> rows, double spacing) {
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            foreach (TableCellLayout cell in rows[rowIndex].Cells) {
                double height = GetSpanningHeight(rows, rowIndex, cell.RowSpan, spacing);
                double slack = Math.Max(0D, height - cell.Style.VerticalInsets - cell.Inline.Height);
                double offset = cell.Style.TableVerticalAlignment == "bottom" ? slack
                    : cell.Style.TableVerticalAlignment == "middle" ? slack / 2D
                    : IsBaselineTableCell(cell) ? Math.Max(0D, rows[rowIndex].Baseline - ResolveTableCellBaseline(cell)) : 0D;
                cell.ContentOffsetY = cell.Style.BorderTopWidth + cell.Style.PaddingTop + offset;
            }
        }
    }

    private static bool IsBaselineTableCell(TableCellLayout cell) =>
        cell.Style.TableVerticalAlignment != "top" && cell.Style.TableVerticalAlignment != "middle" && cell.Style.TableVerticalAlignment != "bottom";

    private static double ResolveTableCellBaseline(TableCellLayout cell) =>
        cell.Style.BorderTopWidth + cell.Style.PaddingTop + (FindTableContentBaseline(cell.Inline.Visuals) ?? cell.Inline.Height);

    private static double? FindTableContentBaseline(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            // Use the same positioned-text baseline as the shared drawing owner.
            if (visual is HtmlRenderText text) return text.Y + text.Font.Size;
            IReadOnlyList<HtmlRenderVisual>? children = GetGroupChildren(visual);
            if (children == null) continue;
            double? baseline = FindTableContentBaseline(children);
            if (baseline.HasValue) return baseline;
        }
        return null;
    }
}
