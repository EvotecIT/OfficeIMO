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
        PdfCore.PdfTableStyle style, double width, double height, int slideNumber, PowerPointToPdfOptions options) {
        Func<string?, double, string?, OfficeFontStyle, double> measure = options.NativeTextMeasure
            ?? OfficeDrawingTextLayout.CreateMetrics(null).MeasureText;
        double gap = style.CellSpacing;
        if (width <= Math.Max(0, table.Columns - 1) * gap || height <= Math.Max(0, table.Rows - 1) * gap) return height;
        double[] widths = NormalizeTableLayoutSegments(style.ColumnWidthPoints!, width - Math.Max(0, table.Columns - 1) * gap);
        double[] heights = NormalizeTableLayoutSegments(style.RowMinHeights!, height - Math.Max(0, table.Rows - 1) * gap);
        bool reportedFit = false;
        for (int row = 0; row < rows.Count; row++) {
            int cellIndex = 0;
            for (int column = 0; column < table.Columns; column++) {
                PptCore.PowerPointTableCell source = table.GetCell(row, column);
                if (source.IsMergedCell) continue;
                options.CancellationToken.ThrowIfCancellationRequested();
                PdfCore.PdfTableCell cell = rows[row][cellIndex];
                int columnEnd = Math.Min(table.Columns, column + cell.ColumnSpan);
                int rowEnd = Math.Min(table.Rows, row + cell.RowSpan);
                double cellWidth = gap * Math.Max(0, columnEnd - column - 1);
                for (int index = column; index < columnEnd; index++) cellWidth += widths[index];
                PdfCore.PdfCellPadding? padding = null;
                style.CellPaddings?.TryGetValue((row, column), out padding);
                double left = padding?.Left ?? style.CellPaddingLeft ?? style.CellPaddingX;
                double right = padding?.Right ?? style.CellPaddingRight ?? style.CellPaddingX;
                double top = padding?.Top ?? style.CellPaddingTop ?? style.CellPaddingY;
                double bottom = padding?.Bottom ?? style.CellPaddingBottom ?? style.CellPaddingY;
                // Leave the original diagnostic path active for impossible margins.
                double innerWidth = Math.Max(1D, cellWidth - left - right);
                double fontSize = row < style.HeaderRowCount ? style.HeaderFontSize ?? style.FontSize ?? 18D
                    : row >= table.Rows - style.FooterRowCount ? style.FooterFontSize ?? style.FontSize ?? 18D
                    : style.FontSize ?? 18D;
                bool rowBold = row < style.HeaderRowCount ? style.HeaderBold : row >= table.Rows - style.FooterRowCount && style.FooterBold;
                var runs = cell.Runs.Select(run => new OfficeRichTextRun(run.Text,
                    run.FontSize ?? fontSize, OfficeColor.Black, run.Bold || rowBold, run.Italic,
                    fontFamily: run.FontFamily ?? options.NativeTextDefaultFontFamily
                        ?? PdfCore.PdfFontNames.ToBaseFontName(run.Font ?? options.PdfOptions?.DefaultFont ?? PdfCore.PdfStandardFont.Helvetica),
                    baseline: run.Baseline == PdfCore.PdfTextBaseline.Superscript ? OfficeTextBaseline.Superscript
                        : run.Baseline == PdfCore.PdfTextBaseline.Subscript ? OfficeTextBaseline.Subscript : OfficeTextBaseline.Normal)).ToArray();
                string text = string.Concat(runs.Select(run => run.Text));
                // An unbreakable label such as Status should remain a word. Use
                // the PDF owner's existing cell fit, rather than inserting breaks.
                bool fittedLabel = false;
                if (!string.IsNullOrWhiteSpace(text) && !text.Any(char.IsWhiteSpace)) {
                    double textWidth = runs.Sum(run => measure(run.Text, run.FontSize, run.FontFamily, run.FontStyle));
                    double minimumSize = style.MinimumShrinkFontSize ?? 6D;
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
                // Keep the authored sizes while measuring: this is conservative
                // for fitted labels and leaves enough room for mixed-size runs.
                double lineHeightFactor = style.LineHeight ?? 1.4D;
                OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutStyledRichTextBlock(runs,
                    innerWidth, double.MaxValue, lineHeightFactor, measure, wrap: !fittedLabel,
                    cancellationToken: options.CancellationToken);
                // PDF table leading is a floor on every line, including lines
                // whose explicit runs are smaller than the row's default font.
                double leading = fontSize * lineHeightFactor;
                double required = Math.Max(leading, layout.Lines.Sum(line => Math.Max(leading, line.LineHeight))) + top + bottom;
                double available = gap * Math.Max(0, rowEnd - row - 1);
                for (int index = row; index < rowEnd; index++) available += heights[index];
                if (required > available) heights[rowEnd - 1] += required - available;
                cellIndex++;
            }
        }
        style.RowMinHeights = heights.Select(value => (double?)value).ToList();
        if (style.FixedRowHeights != null) style.FixedRowHeights = new List<double?>(style.RowMinHeights);
        return heights.Sum() + Math.Max(0, table.Rows - 1) * gap;
    }

    private static double[] NormalizeTableLayoutSegments(IReadOnlyList<double?> segments, double total) {
        double sum = segments.Sum(value => value.GetValueOrDefault());
        return segments.Select(value => sum > 0D ? value.GetValueOrDefault() * total / sum : total / segments.Count).ToArray();
    }
}
