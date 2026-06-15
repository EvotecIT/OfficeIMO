using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void RenderTable(RtfDocument document, RtfTable table, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options, PdfRenderState state) {
        List<PdfCore.PdfTableCell[]> rows = new List<PdfCore.PdfTableCell[]>();
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            RtfTableRow row = table.Rows[rowIndex];
            List<PdfCore.PdfTableCell> cells = new List<PdfCore.PdfTableCell>();
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++) {
                RtfTableCell cell = row.Cells[cellIndex];
                if (cell.HorizontalMerge == RtfTableCellMerge.Continue || cell.VerticalMerge == RtfTableCellMerge.Continue) {
                    continue;
                }

                List<PdfCore.TextRun> runs = BuildCellRuns(document, cell, options, state);
                List<PdfCore.PdfTableCellImage> images = BuildCellImages(cell, options);
                int columnSpan = GetHorizontalMergeSpan(row, cellIndex);
                int rowSpan = GetVerticalMergeSpan(table, rowIndex, cellIndex);
                if (images.Count > 0) {
                    cells.Add(PdfCore.PdfTableCell.WithImages(runs, images, columnSpan: columnSpan, rowSpan: rowSpan));
                } else {
                    cells.Add(PdfCore.PdfTableCell.Merge(runs, columnSpan: columnSpan, rowSpan: rowSpan));
                }
            }

            if (cells.Count > 0) {
                rows.Add(cells.ToArray());
            }
        }

        if (rows.Count > 0) {
            pdf.Table(rows, style: RtfPdfMapping.ToPdfTableStyle(document, table, options));
        }
    }

    private static int GetHorizontalMergeSpan(RtfTableRow row, int cellIndex) {
        if (row.Cells[cellIndex].HorizontalMerge != RtfTableCellMerge.First) {
            return 1;
        }

        int span = 1;
        for (int index = cellIndex + 1; index < row.Cells.Count; index++) {
            if (row.Cells[index].HorizontalMerge != RtfTableCellMerge.Continue) {
                break;
            }

            span++;
        }

        return span;
    }

    private static int GetVerticalMergeSpan(RtfTable table, int rowIndex, int cellIndex) {
        if (table.Rows[rowIndex].Cells[cellIndex].VerticalMerge != RtfTableCellMerge.First) {
            return 1;
        }

        int span = 1;
        for (int index = rowIndex + 1; index < table.Rows.Count; index++) {
            RtfTableRow row = table.Rows[index];
            if (cellIndex >= row.Cells.Count ||
                row.Cells[cellIndex].VerticalMerge != RtfTableCellMerge.Continue) {
                break;
            }

            span++;
        }

        return span;
    }

    private static void RenderImage(RtfImage image, PdfCore.PdfDocument pdf, RtfPdfSaveOptions options) {
        if (!options.IncludeImages || !IsPdfSupportedImage(image) || image.Data.Length == 0) {
            return;
        }

        pdf.Image(image.Data, GetImageWidth(image, options), GetImageHeight(image, options), image.Description);
    }

    private static List<PdfCore.TextRun> BuildCellRuns(RtfDocument document, RtfTableCell cell, RtfPdfSaveOptions options, PdfRenderState state) {
        List<PdfCore.TextRun> runs = new List<PdfCore.TextRun>();
        for (int i = 0; i < cell.Paragraphs.Count; i++) {
            if (i > 0) {
                runs.Add(PdfCore.TextRun.LineBreak());
            }

            AppendParagraphRuns(document, cell.Paragraphs[i], runs, options, state);
        }

        if (runs.Count == 0) {
            runs.Add(PdfCore.TextRun.Normal(string.Empty));
        }

        return runs;
    }

    private static List<PdfCore.PdfTableCellImage> BuildCellImages(RtfTableCell cell, RtfPdfSaveOptions options) {
        List<PdfCore.PdfTableCellImage> images = new List<PdfCore.PdfTableCellImage>();
        if (!options.IncludeImages) {
            return images;
        }

        foreach (RtfParagraph paragraph in cell.Paragraphs) {
            foreach (IRtfInline inline in paragraph.Inlines) {
                if (inline is RtfImage image && IsPdfSupportedImage(image) && image.Data.Length > 0) {
                    images.Add(new PdfCore.PdfTableCellImage(image.Data, GetImageWidth(image, options), GetImageHeight(image, options)));
                }
            }
        }

        return images;
    }

    private static bool IsPdfSupportedImage(RtfImage image) => image.Format == RtfImageFormat.Png || image.Format == RtfImageFormat.Jpeg;

    private static double GetImageWidth(RtfImage image, RtfPdfSaveOptions options) {
        if (image.DesiredWidthTwips.HasValue && image.DesiredWidthTwips.Value > 0) {
            return RtfPdfMapping.TwipsToPoints(image.DesiredWidthTwips.Value);
        }

        return options.DefaultImageWidth;
    }

    private static double GetImageHeight(RtfImage image, RtfPdfSaveOptions options) {
        if (image.DesiredHeightTwips.HasValue && image.DesiredHeightTwips.Value > 0) {
            return RtfPdfMapping.TwipsToPoints(image.DesiredHeightTwips.Value);
        }

        return options.DefaultImageHeight;
    }
}
