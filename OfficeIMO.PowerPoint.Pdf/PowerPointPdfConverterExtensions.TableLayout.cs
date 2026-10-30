using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    // Canvas tables scale their row vector into the supplied frame. Grow both
    // together: increasing RowMinHeights alone gets scaled back and still clips.
    private static double FitTableRows(PptCore.PowerPointTable table, List<PdfCore.PdfTableCell[]> rows,
        PdfCore.PdfTableStyle style, double width, double height, int slideNumber, PowerPointPdfExportOptions options) {
        Func<string?, double, string?, OfficeFontStyle, double> measure = options.NativeTextMeasure
            ?? OfficeDrawingTextLayout.CreateMetrics(null).MeasureText;
        double gap = style.CellSpacing.ToPoints();
        if (width <= Math.Max(0, table.Columns - 1) * gap || height <= Math.Max(0, table.Rows - 1) * gap) return height;
        double[] widths = NormalizeTableLayoutSegments(style.ColumnWidths?.ConvertAll(static length => length?.ToPoints())!, width - Math.Max(0, table.Columns - 1) * gap);
        double[] heights = NormalizeTableLayoutSegments(style.RowMinHeights?.ConvertAll(static length => length?.ToPoints())!, height - Math.Max(0, table.Rows - 1) * gap);
        bool reportedFit = false;
        var cellIndexes = new int[rows.Count];
        foreach (var anchor in table.EnumerateAnchorCells()) {
            int row = anchor.Row, column = anchor.Column;
            int cellIndex = cellIndexes[row]++;
            PptCore.PowerPointTableCell source = anchor.Cell;
            options.CancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfTableCell cell = rows[row][cellIndex];
            int columnEnd = Math.Min(table.Columns, column + cell.ColumnSpan);
            int rowEnd = Math.Min(table.Rows, row + cell.RowSpan);
            double cellWidth = gap * Math.Max(0, columnEnd - column - 1);
            for (int index = column; index < columnEnd; index++) cellWidth += widths[index];
            PdfCore.PdfCellPadding? padding = null;
            style.CellPaddings?.TryGetValue((row, column), out padding);
            double left = padding?.Left?.ToPoints() ?? style.CellPaddingLeft?.ToPoints() ?? style.CellPaddingX.ToPoints();
            double right = padding?.Right?.ToPoints() ?? style.CellPaddingRight?.ToPoints() ?? style.CellPaddingX.ToPoints();
            double top = padding?.Top?.ToPoints() ?? style.CellPaddingTop?.ToPoints() ?? style.CellPaddingY.ToPoints();
            double bottom = padding?.Bottom?.ToPoints() ?? style.CellPaddingBottom?.ToPoints() ?? style.CellPaddingY.ToPoints();
            // Leave the original diagnostic path active for impossible margins.
            double innerWidth = Math.Max(1D, cellWidth - left - right);
            double fontSize = row < style.HeaderRowCount ? style.HeaderFontSize?.ToPoints() ?? style.FontSize?.ToPoints() ?? 18D
                : row >= table.Rows - style.FooterRowCount ? style.FooterFontSize?.ToPoints() ?? style.FontSize?.ToPoints() ?? 18D
                : style.FontSize?.ToPoints() ?? 18D;
            bool rowBold = row < style.HeaderRowCount ? style.HeaderBold : row >= table.Rows - style.FooterRowCount && style.FooterBold;
            OfficeRichTextRun[] runs = CreateTableLayoutRuns(cell.Runs, fontSize, rowBold, options);
            string text = string.Concat(runs.Select(run => run.Text));
            // An unbreakable label such as Status should remain a word. Use
            // the PDF owner's existing cell fit, rather than inserting breaks.
            bool fittedLabel = false;
            if (!string.IsNullOrWhiteSpace(text) && !text.Any(char.IsWhiteSpace)) {
                double textWidth = runs.Sum(run => measure(run.Text, run.FontSize, run.FontFamily, run.FontStyle));
                double minimumSize = style.MinimumShrinkFontSize?.ToPoints() ?? 6D;
                double minimumWidth = runs.Sum(run => measure(run.Text, Math.Min(run.FontSize, minimumSize), run.FontFamily, run.FontStyle));
                if (textWidth > innerWidth + .01D && minimumWidth <= innerWidth) {
                    rows[row][cellIndex] = cell.WithShrinkTextToFit();
                    fittedLabel = true;
                    if (!reportedFit) {
                        AddWarning(options, slideNumber, "table-cell-text-fit-approximation",
                            "PowerPoint table labels wider than their cell are fitted to the cell width to avoid splitting words.");
                        reportedFit = true;
                    }
                }
            }
            double required = MeasureTableParagraphs(cell, fontSize, style.LineHeight ?? 1.4D,
                rowBold, innerWidth, fittedLabel, measure, options) + top + bottom;
            double available = gap * Math.Max(0, rowEnd - row - 1);
            for (int index = row; index < rowEnd; index++) available += heights[index];
            if (required > available) heights[rowEnd - 1] += required - available;
        }
        style.RowMinHeights = (heights.Select(value => (double?)value).ToList()) is { } lengths4814 ? System.Linq.Enumerable.ToList(System.Linq.Enumerable.Select(lengths4814, static length => length.HasValue ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(length.Value) : null)) : null;
        if (style.FixedRowHeights != null) style.FixedRowHeights = style.RowMinHeights;
        return heights.Sum() + Math.Max(0, table.Rows - 1) * gap;
    }

    private static OfficeRichTextRun[] CreateTableLayoutRuns(IEnumerable<PdfCore.PdfTextRun> runs,
        double fontSize, bool rowBold, PowerPointPdfExportOptions options) => runs.Select(run => new OfficeRichTextRun(run.Text,
            (run.FontSize?.ToPoints()) ?? fontSize, OfficeColor.Black, run.Bold || rowBold, run.Italic,
            fontFamily: run.FontFamily ?? options.NativeTextDefaultFontFamily
                ?? PdfCore.PdfFontNames.ToBaseFontName(run.Font ?? options.PdfOptions?.DefaultFont ?? PdfCore.PdfStandardFont.Helvetica),
            baseline: run.Baseline == PdfCore.PdfTextBaseline.Superscript ? OfficeTextBaseline.Superscript
                : run.Baseline == PdfCore.PdfTextBaseline.Subscript ? OfficeTextBaseline.Subscript : OfficeTextBaseline.Normal)).ToArray();

    private static double MeasureTableParagraphs(PdfCore.PdfTableCell cell, double fontSize, double defaultFactor,
        bool rowBold, double width, bool fittedLabel, Func<string?, double, string?, OfficeFontStyle, double> measure,
        PowerPointPdfExportOptions options) {
        IReadOnlyList<PdfCore.PdfTableCellParagraph> paragraphs = cell.Paragraphs.Count > 0
            ? cell.Paragraphs : new[] { new PdfCore.PdfTableCellParagraph(cell.Runs, fontSize: fontSize) };
        double height = 0D;
        double paintedBottom = 0D;
        foreach (PdfCore.PdfTableCellParagraph paragraph in paragraphs) {
            options.CancellationToken.ThrowIfCancellationRequested();
            double size = paragraph.FontSize ?? fontSize;
            bool exact = paragraph.LineSpacing?.Rule == PdfCore.PdfLineSpacingRule.Exact;
            double factor = exact ? paragraph.LineSpacing!.Height!.Value.ToPoints() / size : (paragraph.LineSpacing?.Height?.ToPoints() ?? paragraph.LineSpacing?.MultipleValue) ?? paragraph.LineHeight ?? defaultFactor;
            double leading = size * factor;
            OfficeRichTextRun[] runs = CreateTableLayoutRuns(paragraph.Runs, size, rowBold, options);
            // PDF keeps a hanging indent relative to the paragraph's left
            // margin. Drawing layout instead takes nonnegative offsets from
            // the cell frame for both the first and continuation lines.
            double firstOffset = Math.Max(0D, paragraph.LeftIndent + paragraph.FirstLineIndent);
            double followingOffset = Math.Max(0D, paragraph.LeftIndent);
            var layout = OfficeTextLayoutEngine.LayoutStyledRichTextBlock(runs,
                Math.Max(1D, width - paragraph.RightIndent), double.MaxValue, factor, measure,
                wrap: !fittedLabel, paragraphIndent: new OfficeTextParagraphIndent(firstOffset, followingOffset),
                cancellationToken: options.CancellationToken);
            height += paragraph.SpacingBefore;
            double advance = exact ? Math.Max(1, layout.Lines.Count) * leading
                : Math.Max(leading, layout.Lines.Sum(line => Math.Max(leading, line.LineHeight)));
            // A condensed final line can paint below its advance; grow the row
            // for that ink without changing the paragraph's requested leading.
            paintedBottom = Math.Max(paintedBottom, height + Math.Max(advance, OfficeDrawingTextLayout.RequiredFrameHeight(layout)));
            height += advance + paragraph.SpacingAfter;
        }
        return Math.Max(height, paintedBottom);
    }

    private static double[] NormalizeTableLayoutSegments(IReadOnlyList<double?> segments, double total) {
        double sum = segments.Sum(value => value.GetValueOrDefault());
        return segments.Select(value => sum > 0D ? value.GetValueOrDefault() * total / sum : total / segments.Count).ToArray();
    }
}
