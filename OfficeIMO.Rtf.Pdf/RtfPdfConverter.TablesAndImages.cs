using PdfCore = OfficeIMO.Pdf;
using OfficeIMO.Drawing;

namespace OfficeIMO.Rtf.Pdf;

internal static partial class RtfPdfConverter {
    private static void RenderTable(RtfDocument document, RtfTable table, PdfCore.PdfDocument pdf, RtfToPdfOptions options, PdfRenderState state) {
        List<PdfCore.PdfTableCell[]> rows = new List<PdfCore.PdfTableCell[]>();
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            RtfTableRow row = table.Rows[rowIndex];
            List<PdfCore.PdfTableCell> cells = new List<PdfCore.PdfTableCell>();
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++) {
                RtfTableCell cell = row.Cells[cellIndex];
                if (cell.HorizontalMerge == RtfTableCellMerge.Continue || cell.VerticalMerge == RtfTableCellMerge.Continue) {
                    continue;
                }

                List<PdfCore.PdfTextRun> runs = BuildCellRuns(document, cell, options, state);
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

    private static void RenderImage(RtfImage image, PdfCore.PdfDocument pdf, RtfToPdfOptions options) {
        if (!TryGetRenderableImage(image, options, "Image", out byte[] imageBytes)) {
            return;
        }

        RtfImageLayout layout = image.ResolveLayout(options.DefaultImageWidth * 20d, options.DefaultImageHeight * 20d);
        pdf.Image(imageBytes, layout.VisibleWidthTwips!.Value / 20d, layout.VisibleHeightTwips!.Value / 20d, style: GetImageStyle(image, layout));
    }

    private static List<PdfCore.PdfTextRun> BuildCellRuns(RtfDocument document, RtfTableCell cell, RtfToPdfOptions options, PdfRenderState state) {
        List<PdfCore.PdfTextRun> runs = new List<PdfCore.PdfTextRun>();
        int blockIndex = 0;
        foreach (IRtfBlock block in cell.Blocks) {
            if (blockIndex > 0) {
                runs.Add(PdfCore.PdfTextRun.LineBreak());
            }

            if (block is RtfParagraph paragraph) {
                AppendParagraphRuns(document, paragraph, runs, options, state);
            } else if (block is RtfTable nestedTable) {
                runs.Add(PdfCore.PdfTextRun.Normal(FlattenNestedTableText(nestedTable)));
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
            runs.Add(PdfCore.PdfTextRun.Normal(string.Empty));
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

    private static List<PdfCore.PdfTableCellImage> BuildCellImages(RtfTableCell cell, RtfToPdfOptions options) {
        List<PdfCore.PdfTableCellImage> images = new List<PdfCore.PdfTableCellImage>();
        foreach (RtfParagraph paragraph in cell.Paragraphs) {
            foreach (IRtfInline inline in paragraph.Inlines) {
                if (inline is RtfImage image && TryGetRenderableImage(image, options, "TableCell/Image", out byte[] imageBytes)) {
                    RtfImageLayout layout = image.ResolveLayout(options.DefaultImageWidth * 20d, options.DefaultImageHeight * 20d);
                    images.Add(new PdfCore.PdfTableCellImage(imageBytes, layout.VisibleWidthTwips!.Value / 20d, layout.VisibleHeightTwips!.Value / 20d, GetImageStyle(image, layout)));
                }
            }
        }

        return images;
    }

    private static bool TryGetRenderableImage(RtfImage image, RtfToPdfOptions options, string source, out byte[] imageBytes) {
        if (!TryGetImagePayload(image, options, source, out imageBytes)) return false;
        try {
            RtfImageLayout layout = image.ResolveLayout(options.DefaultImageWidth * 20d, options.DefaultImageHeight * 20d);
            if (!HasNegativeImageCrop(image)) return true;
            var decodeOptions = new OfficeRasterDecodeOptions { CancellationToken = options.CancellationToken };
            if (!OfficeRasterImageDecoder.TryDecode(imageBytes, decodeOptions, out OfficeRasterImage? raster, out _) || raster == null) {
                AddConversionWarning(options, "ImageCropDecodeFailed", source, "The picture could not be decoded to preserve crop padding.", RtfConversionAction.Blocked);
                return false;
            }
            double ratioX = raster.Width / layout.WidthTwips!.Value;
            double ratioY = raster.Height / layout.HeightTwips!.Value;
            double width = Math.Ceiling(layout.VisibleWidthTwips!.Value / layout.ScaleX * ratioX);
            double height = Math.Ceiling(layout.VisibleHeightTwips!.Value / layout.ScaleY * ratioY);
            if (width > int.MaxValue || height > int.MaxValue || width * height > decodeOptions.MaximumDecodedPixels) {
                AddConversionWarning(options, "ImageCropBudgetExceeded", source, "Picture crop padding exceeds the shared raster pixel limit.", RtfConversionAction.Blocked);
                return false;
            }
            OfficeRasterImage padded = OfficeImageComposer.ComposeRaster((int)width, (int)height, OfficeColor.Transparent,
                new[] { OfficeImageLayer.FromRaster(raster, -(image.CropLeftTwips ?? 0) * ratioX, -(image.CropTopTwips ?? 0) * ratioY, raster.Width, raster.Height) },
                beforeLayers: null, afterLayers: null, fonts: null, cancellationToken: options.CancellationToken);
            using var output = new MemoryStream();
            OfficePngWriter.EncodeTo(padded, output, new OfficePngEncodeOptions(), options.CancellationToken);
            imageBytes = output.ToArray();
            ReportImageSubstitution(options, source, image.Format, "PNG", "Picture crop padding was composed through the shared raster engine.");
            return true;
        } catch (InvalidDataException exception) {
            AddConversionWarning(options, "ImageLayoutInvalid", source, exception.Message, RtfConversionAction.Blocked);
            return false;
        }
    }

    private static bool HasNegativeImageCrop(RtfImage image) => image.CropLeftTwips < 0 || image.CropTopTwips < 0 || image.CropRightTwips < 0 || image.CropBottomTwips < 0;

    private static bool TryGetImagePayload(RtfImage image, RtfToPdfOptions options, string source, out byte[] imageBytes) {
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
            if (converted != null && PdfCore.PdfDocument.TryPrepareImageBytes(
                    converted,
                    out imageBytes,
                    out _,
                    out _,
                    out reason)) {
                ReportImageSubstitution(options, source, image.Format, "PDF raster", "The RTF image was converted by the configured image converter and prepared through OfficeIMO.Drawing.");
                return true;
            }

            AddConversionWarning(
                options,
                "ImageConversionFailed",
                source,
                reason ?? "The configured image converter did not return a raster payload supported by OfficeIMO.Drawing.",
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

    private static void ReportImageSubstitution(RtfToPdfOptions options, string source, RtfImageFormat sourceFormat, string targetFormat, string message) {
        var details = new Dictionary<string, string> {
            ["SourceFormat"] = sourceFormat.ToString(),
            ["TargetFormat"] = targetFormat
        };
        details["RtfAction"] = RtfConversionAction.Substituted.ToString();
        options.Report.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.Rtf.Pdf",
            "ImageConverted",
            source,
            message,
            PdfCore.PdfConversionWarningSeverity.Information,
            details: details));
    }

    private static PdfCore.PdfImageStyle GetImageStyle(RtfImage image, RtfImageLayout layout) {
        var style = new PdfCore.PdfImageStyle { AlternativeText = string.IsNullOrWhiteSpace(image.Description) ? null : image.Description };
        double left = (image.CropLeftTwips ?? 0) / layout.WidthTwips!.Value;
        double top = (image.CropTopTwips ?? 0) / layout.HeightTwips!.Value;
        double right = (image.CropRightTwips ?? 0) / layout.WidthTwips!.Value;
        double bottom = (image.CropBottomTwips ?? 0) / layout.HeightTwips!.Value;
        if (HasNegativeImageCrop(image)) return style;
        if (left != 0 || top != 0 || right != 0 || bottom != 0) style.SourceCrop = new PdfCore.PdfImageSourceCrop(left, top, right, bottom);
        return style;
    }
}
