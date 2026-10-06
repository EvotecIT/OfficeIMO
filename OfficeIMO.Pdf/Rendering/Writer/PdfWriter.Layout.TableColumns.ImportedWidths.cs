namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static double MeasureImportedTableMinimumTextWidth(TableCellLayout cell,
        PdfStandardFont font, double size, PdfOptions? options, double runFontSizeScale = 1D, double minimumShrinkFontSize = 0D) {
        // The normal rich layout resolves the actual run fonts, sizes, shaping
        // and break opportunities. Keep those measurements when finding the
        // minimum width of an imported grid.
        TableCellTextLayout layout = CreateTableCellTextLayout(cell, TableCellNoWrapWidth,
            font, size, size * 1.25D, options, runFontSizeScale, minimumShrinkFontSize);
        double minimum = 0D;
        foreach (var line in layout.Lines) {
            if (cell.NoWrap) {
                minimum = Math.Max(minimum, MeasureRichLineWidth(line, options));
                continue;
            }
            double word = 0D;
            foreach (RichSeg segment in line) {
                if (segment.LeadingSpace || segment.LeadingIsTab) word = 0D;
                word += segment.MeasuredWidth;
                minimum = Math.Max(minimum, word);
                if (segment.EndsWithHardBreak || segment.EndsWithTextSeparator) word = 0D;
            }
        }
        return minimum;
    }

    private static double ResolveImportedTableCellWrapWidth(TableCellLayout cell, double innerWidth,
        PdfStandardFont font, double size, PdfOptions? options, double runFontSizeScale,
        double minimumShrinkFontSize, bool wrapOversizedNoWrap) {
        bool noWrap = cell.NoWrap;
        // Word can enlarge an automatic grid for a no-wrap cell, but falls
        // back to wrapping when its text cannot fit inside the page frame.
        if (noWrap && wrapOversizedNoWrap &&
            MeasureImportedTableMinimumTextWidth(cell, font, size, options, runFontSizeScale, minimumShrinkFontSize) > innerWidth + .001D)
            noWrap = false;
        return GetTableCellWrapWidth(innerWidth, noWrap);
    }
}
