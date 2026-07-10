using PdfCore = OfficeIMO.Pdf;
using OfficeIMO.Drawing;

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
        if (!TryGetRenderableImage(image, options, "Image", out byte[] imageBytes)) {
            return;
        }

        pdf.Image(imageBytes, GetImageWidth(image, options), GetImageHeight(image, options), image.Description);
    }

    private static List<PdfCore.TextRun> BuildCellRuns(RtfDocument document, RtfTableCell cell, RtfPdfSaveOptions options, PdfRenderState state) {
        List<PdfCore.TextRun> runs = new List<PdfCore.TextRun>();
        int blockIndex = 0;
        foreach (IRtfBlock block in cell.Blocks) {
            if (blockIndex > 0) {
                runs.Add(PdfCore.TextRun.LineBreak());
            }

            if (block is RtfParagraph paragraph) {
                AppendParagraphRuns(document, paragraph, runs, options, state);
            } else if (block is RtfTable nestedTable) {
                runs.Add(PdfCore.TextRun.Normal(FlattenNestedTableText(nestedTable)));
                AddConversionWarning(
                    options,
                    "NestedTableFlattened",
                    "TableCell/NestedTable",
                    "A nested RTF table was flattened to delimited text inside its PDF table cell.",
                    RtfConversionAction.Flattened);
            }

            blockIndex++;
        }

        if (runs.Count == 0) {
            runs.Add(PdfCore.TextRun.Normal(string.Empty));
        }

        return runs;
    }

    private static string FlattenNestedTableText(RtfTable table) {
        return string.Join(" / ", table.Rows.Select(row =>
            string.Join(" | ", row.Cells.Select(cell =>
                string.Join(" ", cell.Blocks.Select(block => block is RtfParagraph paragraph
                    ? paragraph.ToPlainText()
                    : block is RtfTable nested ? FlattenNestedTableText(nested) : string.Empty)
                    .Where(text => !string.IsNullOrWhiteSpace(text)))))));
    }

    private static List<PdfCore.PdfTableCellImage> BuildCellImages(RtfTableCell cell, RtfPdfSaveOptions options) {
        List<PdfCore.PdfTableCellImage> images = new List<PdfCore.PdfTableCellImage>();
        foreach (RtfParagraph paragraph in cell.Paragraphs) {
            foreach (IRtfInline inline in paragraph.Inlines) {
                if (inline is RtfImage image && TryGetRenderableImage(image, options, "TableCell/Image", out byte[] imageBytes)) {
                    images.Add(new PdfCore.PdfTableCellImage(imageBytes, GetImageWidth(image, options), GetImageHeight(image, options)));
                }
            }
        }

        return images;
    }

    private static bool TryGetRenderableImage(RtfImage image, RtfPdfSaveOptions options, string source, out byte[] imageBytes) {
        imageBytes = Array.Empty<byte>();
        if (!options.IncludeImages) {
            AddConversionWarning(
                options,
                "ImageSkipped",
                source,
                "An RTF image was skipped because IncludeImages is false.",
                new Dictionary<string, string> {
                    ["Format"] = image.Format.ToString()
                });
            return false;
        }

        if (image.Data.Length == 0) {
            AddConversionWarning(
                options,
                "ImageSkipped",
                source,
                "An RTF image was skipped because it does not contain image data.",
                new Dictionary<string, string> {
                    ["Format"] = image.Format.ToString()
                });
            return false;
        }

        if (image.Format == RtfImageFormat.Png || image.Format == RtfImageFormat.Jpeg) {
            imageBytes = image.Data;
            return true;
        }

        if (image.Format == RtfImageFormat.Dib && OfficeImagePngConverter.TryConvertDibToPng(image.Data, out imageBytes)) {
            ReportImageSubstitution(options, source, image.Format, "PNG", "The RTF DIB image was converted to PNG through OfficeIMO.Drawing.");
            return true;
        }

        if (options.ImageConverter != null) {
            byte[]? converted = options.ImageConverter(image);
            string? reason = null;
            if (converted != null && OfficeImagePdfCompatibility.TryValidate(converted, out _, out reason)) {
                imageBytes = converted;
                ReportImageSubstitution(options, source, image.Format, "PNG/JPEG", "The RTF image was converted by the configured image converter.");
                return true;
            }

            AddConversionWarning(
                options,
                "ImageConversionFailed",
                source,
                reason ?? "The configured image converter did not return PNG or JPEG bytes.",
                new Dictionary<string, string> {
                    ["Format"] = image.Format.ToString()
                });
            return false;
        }

        AddConversionWarning(
            options,
            "UnsupportedImage",
            source,
            "Only PNG and JPEG RTF images can be embedded directly in PDF output.",
            new Dictionary<string, string> {
                ["Format"] = image.Format.ToString()
            });
        return false;
    }

    private static void ReportImageSubstitution(RtfPdfSaveOptions options, string source, RtfImageFormat sourceFormat, string targetFormat, string message) {
        var details = new Dictionary<string, string> {
            ["SourceFormat"] = sourceFormat.ToString(),
            ["TargetFormat"] = targetFormat
        };
        options.RtfConversionReport.Add(
            RtfConversionSeverity.Information,
            "ImageConverted",
            message,
            RtfConversionAction.Substituted,
            sourcePath: source,
            feature: "Image",
            detail: string.Join(";", details.Select(pair => pair.Key + "=" + pair.Value)));
    }

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
