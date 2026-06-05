using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private const double TableDiagnosticDefaultFontSize = 11D;
    private const double TableDiagnosticDefaultPaddingX = 4D;
    private const double TableDiagnosticDefaultPaddingY = 2D;

    private static void AddLayoutWarning(
        PowerPointPdfSaveOptions options,
        int slideNumber,
        string code,
        string message,
        PdfCore.PdfLayoutDiagnosticKind kind,
        string source,
        string diagnosticMessage,
        double x,
        double y,
        double width,
        double height) {
        AddWarning(
            options,
            slideNumber,
            code,
            message,
            new PdfCore.PdfLayoutDiagnostic(kind, source, diagnosticMessage, x, y, width, height));
    }

    private static void AddPowerPointTableLayoutDiagnostics(PowerPointPdfSaveOptions options, int slideNumber, PptCore.PowerPointTable table, double x, double y, double width, double height) {
        List<double?> columnWidths = CreateColumnWidths(table, width);
        List<double?> rowHeights = CreateRowHeights(table, height);
        bool added = false;

        ForEachTableAnchorCell(table, (row, column, cell) => {
            if (added || row >= rowHeights.Count || column >= columnWidths.Count) {
                return;
            }

            (int rowSpan, int columnSpan) = cell.Merge;
            double cellWidth = SumDimensions(columnWidths, column, Math.Max(1, columnSpan));
            double cellHeight = SumDimensions(rowHeights, row, Math.Max(1, rowSpan));
            if (cellWidth <= 0D || cellHeight <= 0D) {
                return;
            }

            double innerWidth = Math.Max(1D, cellWidth - GetCellPaddingLeft(cell) - GetCellPaddingRight(cell));
            double innerHeight = Math.Max(0D, cellHeight - GetCellPaddingTop(cell) - GetCellPaddingBottom(cell));
            double fontSize = ResolveTableCellDiagnosticFontSize(cell);
            double estimatedHeight = EstimateTableCellTextHeight(CreatePdfTableCellRuns(cell), innerWidth, fontSize);
            if (estimatedHeight <= innerHeight + 0.5D) {
                return;
            }

            double cellX = x + SumDimensions(columnWidths, 0, column);
            double cellY = y + SumDimensions(rowHeights, 0, row);
            AddLayoutWarning(
                options,
                slideNumber,
                "table-cell-overflow",
                "PowerPoint table cell text may be clipped because the mapped PDF cell is too small for its content.",
                PdfCore.PdfLayoutDiagnosticKind.ClippedContent,
                "PowerPointTableCell",
                "The mapped PDF table cell is likely too small for all PowerPoint cell text.",
                cellX,
                cellY,
                cellWidth,
                cellHeight);
            added = true;
        });
    }

    private static void AddPowerPointListLayoutDiagnostics(PowerPointPdfSaveOptions options, int slideNumber, PptCore.PowerPointTextBox textBox, double x, double y, double width, double height) {
        foreach (PptCore.PowerPointParagraph paragraph in textBox.Paragraphs) {
            if (!HasListMarker(paragraph)) {
                continue;
            }

            if (!paragraph.LeftMarginPoints.HasValue && !paragraph.IndentPoints.HasValue) {
                continue;
            }

            AddLayoutWarning(
                options,
                slideNumber,
                "list-indent-simplified",
                "Rendered a PowerPoint list using PDF text prefixes because explicit PowerPoint list indentation is not yet mapped to PDF hanging-indent layout.",
                PdfCore.PdfLayoutDiagnosticKind.SimplifiedContent,
                "PowerPointList",
                "Explicit PowerPoint list indentation was simplified to a PDF text prefix.",
                x,
                y,
                width,
                height);
            return;
        }
    }

    private static void AddPowerPointPictureAspectRatioDiagnostic(
        PowerPointPdfSaveOptions options,
        int slideNumber,
        byte[] imageBytes,
        PptCore.PowerPointPictureCrop crop,
        OfficeImageFit fit,
        double x,
        double y,
        double width,
        double height) {
        if (!options.WarnOnPictureAspectRatioDistortion ||
            fit != OfficeImageFit.Stretch ||
            crop.HasCrop ||
            width <= 0D ||
            height <= 0D ||
            !OfficeImageReader.TryIdentify(imageBytes, fileName: null, out OfficeImageInfo imageInfo) ||
            imageInfo.Width <= 0 ||
            imageInfo.Height <= 0) {
            return;
        }

        double imageAspect = imageInfo.Width / (double)imageInfo.Height;
        double targetAspect = width / height;
        double ratio = Math.Max(imageAspect, targetAspect) / Math.Min(imageAspect, targetAspect);
        if (ratio <= 1.02D) {
            return;
        }

        AddLayoutWarning(
            options,
            slideNumber,
            "picture-aspect-distortion",
            "Rendered an uncropped PowerPoint picture with Stretch fit into a frame whose aspect ratio differs from the source image. Set PowerPointPdfSaveOptions.PictureFit to Contain or Cover to preserve aspect ratio.",
            PdfCore.PdfLayoutDiagnosticKind.SimplifiedContent,
            "PowerPointPicture",
            "The mapped PDF picture frame can distort the source image aspect ratio.",
            x,
            y,
            width,
            height);
    }

    private static double SumDimensions(IReadOnlyList<double?> values, int start, int count) {
        double total = 0D;
        int end = Math.Min(values.Count, start + count);
        for (int index = Math.Max(0, start); index < end; index++) {
            total += values[index].GetValueOrDefault();
        }

        return total;
    }

    private static double EstimateTableCellTextHeight(IReadOnlyList<PdfCore.TextRun> runs, double innerWidth, double fallbackFontSize) {
        int lineCount = 1;
        double lineWidth = 0D;
        double maxLineHeight = fallbackFontSize * 1.2D;

        foreach (PdfCore.TextRun run in runs) {
            if (run.Text == "\n") {
                lineCount++;
                lineWidth = 0D;
                continue;
            }

            double fontSize = run.FontSize.GetValueOrDefault(fallbackFontSize);
            maxLineHeight = Math.Max(maxLineHeight, fontSize * 1.2D);
            foreach (string token in SplitTableDiagnosticTokens(run.Text)) {
                if (token == "\n") {
                    lineCount++;
                    lineWidth = 0D;
                    continue;
                }

                double tokenWidth = EstimateDiagnosticTextWidth(token, fontSize, run.Bold);
                if (lineWidth > 0D && lineWidth + tokenWidth > innerWidth) {
                    lineCount++;
                    lineWidth = 0D;
                }

                if (tokenWidth > innerWidth) {
                    int extraLines = Math.Max(0, (int)Math.Ceiling(tokenWidth / innerWidth) - 1);
                    lineCount += extraLines;
                    lineWidth = tokenWidth - (extraLines * innerWidth);
                } else {
                    lineWidth += tokenWidth;
                }
            }
        }

        return lineCount * maxLineHeight;
    }

    private static IEnumerable<string> SplitTableDiagnosticTokens(string text) {
        if (string.IsNullOrEmpty(text)) {
            yield break;
        }

        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int line = 0; line < lines.Length; line++) {
            if (line > 0) {
                yield return "\n";
            }

            string[] parts = lines[line].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            for (int index = 0; index < parts.Length; index++) {
                yield return index == 0 ? parts[index] : " " + parts[index];
            }
        }
    }

    private static double EstimateDiagnosticTextWidth(string text, double fontSize, bool bold) {
        if (text == "\n") {
            return 0D;
        }

        double widthFactor = bold ? 0.57D : 0.53D;
        return text.Length * fontSize * widthFactor;
    }

    private static double ResolveTableCellDiagnosticFontSize(PptCore.PowerPointTableCell cell) {
        int? fontSize = cell.FontSize;
        if (fontSize.HasValue && fontSize.Value > 0) {
            return fontSize.Value;
        }

        return TableDiagnosticDefaultFontSize;
    }

    private static double GetCellPaddingLeft(PptCore.PowerPointTableCell cell) =>
        cell.PaddingLeftPoints.GetValueOrDefault(TableDiagnosticDefaultPaddingX);

    private static double GetCellPaddingRight(PptCore.PowerPointTableCell cell) =>
        cell.PaddingRightPoints.GetValueOrDefault(TableDiagnosticDefaultPaddingX);

    private static double GetCellPaddingTop(PptCore.PowerPointTableCell cell) =>
        cell.PaddingTopPoints.GetValueOrDefault(TableDiagnosticDefaultPaddingY);

    private static double GetCellPaddingBottom(PptCore.PowerPointTableCell cell) =>
        cell.PaddingBottomPoints.GetValueOrDefault(TableDiagnosticDefaultPaddingY);
}
