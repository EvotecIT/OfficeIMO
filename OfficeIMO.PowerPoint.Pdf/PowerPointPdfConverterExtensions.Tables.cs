using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderTable(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTable table, double x, double y, double width, double height, int slideNumber, PowerPointPdfSaveOptions options) {
        if (table.Rows == 0 || table.Columns == 0) {
            AddWarning(options, slideNumber, "empty-table", "Skipped an empty PowerPoint table.");
            return;
        }

        var rows = new List<PdfCore.PdfTableCell[]>();
        for (int rowIndex = 0; rowIndex < table.Rows; rowIndex++) {
            var pdfCells = new List<PdfCore.PdfTableCell>();
            for (int columnIndex = 0; columnIndex < table.Columns; columnIndex++) {
                PptCore.PowerPointTableCell cell = table.GetCell(rowIndex, columnIndex);
                if (cell.IsMergedCell) {
                    continue;
                }

                pdfCells.Add(CreatePdfTableCell(cell));
            }

            rows.Add(pdfCells.ToArray());
        }

        PdfCore.PdfTableStyle style = CreateTableStyle(table);
        bool reportedCellOverflow = false;
        try {
            canvas.Table(rows, x, y, width, height, style, table.Rotation ?? 0D, diagnostic => {
                if (reportedCellOverflow || diagnostic.Kind != PdfCore.PdfLayoutDiagnosticKind.ClippedContent) {
                    return;
                }

                reportedCellOverflow = true;
                AddWarning(
                    options,
                    slideNumber,
                    "table-cell-overflow",
                    "PowerPoint table cell text was clipped because the PDF table render pass found more text than fits a mapped cell.",
                    new PdfCore.PdfLayoutDiagnostic(
                        PdfCore.PdfLayoutDiagnosticKind.ClippedContent,
                        "PowerPointTableCell",
                        diagnostic.Message,
                        diagnostic.X ?? x,
                        diagnostic.Y ?? y,
                        diagnostic.Width ?? width,
                        diagnostic.Height ?? height));
            });
        } catch (Exception ex) {
            AddLayoutWarning(
                options,
                slideNumber,
                "unsupported-table",
                "Skipped a PowerPoint table because it could not be rendered as a PDF table: " + ex.Message,
                PdfCore.PdfLayoutDiagnosticKind.SkippedContent,
                "PowerPointTable",
                "The PowerPoint table could not be rendered by the fixed-position PDF table renderer.",
                x,
                y,
                width,
                height);
        }
    }

    private static PdfCore.PdfTableCell CreatePdfTableCell(PptCore.PowerPointTableCell cell) {
        (int rowSpan, int columnSpan) = cell.Merge;
        return PdfCore.PdfTableCell.Merge(CreatePdfTableCellRuns(cell), Math.Max(1, columnSpan), Math.Max(1, rowSpan));
    }

    private static IReadOnlyList<PdfCore.TextRun> CreatePdfTableCellRuns(PptCore.PowerPointTableCell cell) {
        var runs = new List<PdfCore.TextRun>();
        A.TextBody? textBody = cell.Cell.TextBody;
        if (textBody != null) {
            bool hasParagraph = false;
            foreach (A.Paragraph paragraph in textBody.Elements<A.Paragraph>()) {
                if (hasParagraph) {
                    runs.Add(PdfCore.TextRun.LineBreak());
                }

                AppendPdfTableCellParagraphRuns(runs, paragraph, cell);
                hasParagraph = true;
            }
        }

        if (runs.Count == 0) {
            runs.Add(CreatePdfTableCellTextRun(cell, cell.Text ?? string.Empty));
        }

        return runs;
    }

    private static void AppendPdfTableCellParagraphRuns(List<PdfCore.TextRun> runs, A.Paragraph paragraph, PptCore.PowerPointTableCell cell) {
        foreach (OpenXmlElement child in paragraph.ChildElements) {
            switch (child) {
                case A.Run run:
                    foreach (A.Text text in run.Elements<A.Text>()) {
                        runs.Add(CreatePdfTableCellTextRun(cell, run, text.Text ?? string.Empty));
                    }

                    break;
                case A.Break:
                    runs.Add(PdfCore.TextRun.LineBreak());
                    break;
                case A.Field field:
                    string fieldText = field.Text?.Text ?? field.InnerText ?? string.Empty;
                    if (!string.IsNullOrEmpty(fieldText)) {
                        runs.Add(CreatePdfTableCellTextRun(cell, fieldText));
                    }

                    break;
            }
        }
    }

    private static PdfCore.TextRun CreatePdfTableCellTextRun(PptCore.PowerPointTableCell cell, string text) {
        return new PdfCore.TextRun(
            text,
            bold: cell.Bold,
            italic: cell.Italic,
            color: ParsePdfColor(cell.Color),
            fontSize: cell.FontSize,
            font: MapFont(cell.FontName));
    }

    private static PdfCore.PdfTableStyle CreateTableStyle(PptCore.PowerPointTable table) {
        var style = PdfCore.TableStyles.Light().Clone();
        style.HeaderRowCount = table.FirstRow ? 1 : 0;
        style.FooterRowCount = table.LastRow ? 1 : 0;
        style.RowStripeFill = table.BandedRows ? style.RowStripeFill : null;
        style.ColumnWidthPoints = CreateColumnWidths(table, table.WidthPoints);
        style.RowMinHeights = CreateRowHeights(table, table.HeightPoints);
        style.CellFills = CreateTableCellFills(table);
        style.CellPaddings = CreateTableCellPaddings(table);
        style.CellAlignments = CreateTableCellAlignments(table);
        style.CellVerticalAlignments = CreateTableCellVerticalAlignments(table);
        style.CellBorders = CreateTableCellBorders(table);
        return style;
    }

    private static List<double?> CreateColumnWidths(PptCore.PowerPointTable table, double tableWidth) {
        var widths = new List<double?>(table.Columns);
        double fallbackWidth = table.Columns > 0 ? tableWidth / table.Columns : tableWidth;
        for (int column = 0; column < table.Columns; column++) {
            double width = table.GetColumnWidthPoints(column);
            widths.Add(width > 0D ? width : fallbackWidth);
        }

        return widths;
    }

    private static List<double?> CreateRowHeights(PptCore.PowerPointTable table, double tableHeight) {
        var heights = new List<double?>(table.Rows);
        double fallbackHeight = table.Rows > 0 ? tableHeight / table.Rows : tableHeight;
        for (int row = 0; row < table.Rows; row++) {
            double height = table.GetRowHeightPoints(row);
            heights.Add(height > 0D ? height : fallbackHeight);
        }

        return heights;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfColor>? CreateTableCellFills(PptCore.PowerPointTable table) {
        var fills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            PdfCore.PdfColor? fill = ParsePdfColor(cell.FillColor);
            if (fill.HasValue) {
                fills[(row, column)] = fill.Value;
            }
        });

        return fills.Count == 0 ? null : fills;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>? CreateTableCellPaddings(PptCore.PowerPointTable table) {
        var paddings = new Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            if (cell.PaddingLeftPoints.HasValue || cell.PaddingRightPoints.HasValue || cell.PaddingTopPoints.HasValue || cell.PaddingBottomPoints.HasValue) {
                paddings[(row, column)] = new PdfCore.PdfCellPadding {
                    Left = cell.PaddingLeftPoints,
                    Right = cell.PaddingRightPoints,
                    Top = cell.PaddingTopPoints,
                    Bottom = cell.PaddingBottomPoints
                };
            }
        });

        return paddings.Count == 0 ? null : paddings;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>? CreateTableCellAlignments(PptCore.PowerPointTable table) {
        var alignments = new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            PdfCore.PdfColumnAlign? align = MapColumnAlign(cell.HorizontalAlignment);
            if (align.HasValue) {
                alignments[(row, column)] = align.Value;
            }
        });

        return alignments.Count == 0 ? null : alignments;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>? CreateTableCellVerticalAlignments(PptCore.PowerPointTable table) {
        var alignments = new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            PdfCore.PdfCellVerticalAlign? align = MapVerticalAlign(cell.VerticalAlignment);
            if (align.HasValue) {
                alignments[(row, column)] = align.Value;
            }
        });

        return alignments.Count == 0 ? null : alignments;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? CreateTableCellBorders(PptCore.PowerPointTable table) {
        var borders = new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            PdfCore.PdfColor? borderColor = ParsePdfColor(cell.BorderColor);
            if (borderColor.HasValue) {
                borders[(row, column)] = new PdfCore.PdfCellBorder {
                    Color = borderColor.Value,
                    Width = 0.75D
                };
            }
        });

        return borders.Count == 0 ? null : borders;
    }

    private static void ForEachTableAnchorCell(PptCore.PowerPointTable table, Action<int, int, PptCore.PowerPointTableCell> action) {
        for (int row = 0; row < table.Rows; row++) {
            for (int column = 0; column < table.Columns; column++) {
                PptCore.PowerPointTableCell cell = table.GetCell(row, column);
                if (!cell.IsMergedCell) {
                    action(row, column, cell);
                }
            }
        }
    }
}
