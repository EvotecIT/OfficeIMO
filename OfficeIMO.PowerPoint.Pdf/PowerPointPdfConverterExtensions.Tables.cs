using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderTable(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTable table, double x, double y, double width, double height, int slideNumber, PowerPointPdfExportOptions options,
        double? slideWidth = null, double? slideHeight = null, bool hasAncestorPageClip = false) {
        if (table.Rows == 0 || table.Columns == 0) {
            AddWarning(options, slideNumber, "empty-table", "Skipped an empty PowerPoint table.", lossKind: OfficeConversionLossKind.Omission);
            return;
        }

        // PowerPoint draws a table at its grid size (column widths, row heights growing with content), not at
        // its graphic-frame size. Map the grid through the frame's scale so group transforms still apply.
        (width, height) = ResolveTableGridSize(table, width, height);
        string? fallbackFontFamily = string.IsNullOrWhiteSpace(options.FontFamily)
            ? ResolveTableDefaultFontFamily(table)
            : null;
        PdfCore.PdfTableStyle style = CreateTableStyle(table, options);
        double bodyFontSize = style.FontSize ?? (options.PdfOptions ?? new PdfCore.PdfCreateOptions()).DefaultFontSize;
        var rowCells = Enumerable.Range(0, table.Rows).Select(_ => new List<PdfCore.PdfTableCell>()).ToArray();
        foreach (var anchor in table.EnumerateAnchorCells()) {
            int rowIndex = anchor.Row;
            double fallbackFontSize = rowIndex < style.HeaderRowCount ? style.HeaderFontSize ?? bodyFontSize
                : rowIndex >= table.Rows - style.FooterRowCount ? style.FooterFontSize ?? bodyFontSize : bodyFontSize;
            IReadOnlyList<A.TextCharacterPropertiesType> tableDefaults = options.PdfOptions?.HasExplicitDefaultTableStyle == true
                ? Array.Empty<A.TextCharacterPropertiesType>()
                : CreateTableTextDefaults(table, rowIndex, anchor.Column);
            rowCells[rowIndex].Add(CreatePdfTableCell(anchor.Cell, fallbackFontFamily, fallbackFontSize,
                style.LineHeight ?? 1.2D, slideNumber, options, tableDefaults));
        }
        var rows = rowCells.Select(row => row.ToArray()).ToList();

        double authoredHeight = height;
        height = FitTableRows(table, rows, style, width, height, slideNumber, options);
        bool reportedCellOverflow = height > authoredHeight + .001D
            && ReportTableSlideOverflow(x, y, width, height, slideWidth, slideHeight, slideNumber, options);
        try {
            Action<PdfCore.PdfPageCanvas> render = target => target.Table(rows, x, y, width, height, style, table.Rotation ?? 0D, diagnostic => {
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
            if (reportedCellOverflow && !hasAncestorPageClip) {
                canvas.Clip(0D, 0D, slideWidth!.Value, slideHeight!.Value, render);
            } else {
                render(canvas);
            }
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

    private static PdfCore.PdfTableCell CreatePdfTableCell(
        PptCore.PowerPointTableCell cell,
        string? fallbackFontFamily,
        double fallbackFontSize,
        double fallbackLineHeight,
        int slideNumber,
        PowerPointPdfExportOptions options,
        IReadOnlyList<A.TextCharacterPropertiesType> tableDefaults) {
        (int rowSpan, int columnSpan) = cell.Merge;
        IReadOnlyList<PdfCore.PdfTableCellParagraph> paragraphs = CreatePdfTableCellParagraphs(cell,
            fallbackFontFamily, fallbackFontSize, fallbackLineHeight, slideNumber, options, tableDefaults);
        var runs = new List<PdfCore.PdfTextRun>();
        foreach (PdfCore.PdfTableCellParagraph paragraph in paragraphs) {
            if (runs.Count > 0) runs.Add(PdfCore.PdfTextRun.LineBreak());
            runs.AddRange(paragraph.Runs);
        }
        if (runs.Count == 0) {
            runs.Add(CreatePdfTableCellTextRun(cell, cell.Text ?? string.Empty, fallbackFontFamily));
        }
        return new PdfCore.PdfTableCell(runs, paragraphs, Math.Max(1, columnSpan), rowSpan: Math.Max(1, rowSpan));
    }

    private static void AppendPdfTableCellParagraphRuns(
        List<PdfCore.PdfTextRun> runs,
        A.Paragraph paragraph,
        PptCore.PowerPointTableCell cell,
        string? fallbackFontFamily,
        double fallbackFontSize,
        int slideNumber,
        PowerPointPdfExportOptions options,
        IReadOnlyList<A.TextCharacterPropertiesType> tableDefaults) {
        foreach (OpenXmlElement child in paragraph.ChildElements) {
            switch (child) {
                case A.Run run:
                    foreach (A.Text text in run.Elements<A.Text>()) {
                        runs.Add(CreatePdfTableCellTextRun(
                            cell, paragraph, run.RunProperties, text.Text ?? string.Empty, fallbackFontFamily, slideNumber, options, tableDefaults, fallbackFontSize));
                    }

                    break;
                case A.Break:
                    runs.Add(PdfCore.PdfTextRun.LineBreak());
                    break;
                case A.Field field:
                    string fieldText = field.Text?.Text ?? field.InnerText ?? string.Empty;
                    if (!string.IsNullOrEmpty(fieldText)) {
                        runs.Add(CreatePdfTableCellTextRun(
                            cell, paragraph, field.RunProperties, fieldText, fallbackFontFamily, slideNumber, options, tableDefaults, fallbackFontSize));
                    }

                    break;
            }
        }
    }

    private static PdfCore.PdfTextRun CreatePdfTableCellTextRun(PptCore.PowerPointTableCell cell, string text, string? fallbackFontFamily) {
        string? fontFamily = cell.FontName ?? fallbackFontFamily;
        return new PdfCore.PdfTextRun(
            text,
            bold: cell.Bold,
            italic: cell.Italic,
            color: ParsePdfColor(cell.Color),
            fontSize: cell.FontSize,
            font: MapFont(fontFamily),
            fontFamily: fontFamily);
    }

    private static PdfCore.PdfTableStyle CreateTableStyle(PptCore.PowerPointTable table, PowerPointPdfExportOptions options) {
        PdfCore.PdfTableStyle style = CreateBaseTableStyle(options);
        bool includePresentationTableStyle = options.PdfOptions?.HasExplicitDefaultTableStyle != true;
        style.HeaderRowCount = table.FirstRow ? 1 : 0;
        style.RepeatHeaderRowCount = style.HeaderRowCount == 0 ? 0 : style.HeaderRowCount;
        style.FooterRowCount = table.LastRow ? 1 : 0;
        style.RowStripeFill = table.BandedRows ? style.RowStripeFill : null;
        style.ColumnWidthPoints = CreateColumnWidths(table, table.WidthPoints);
        style.RowMinHeights = CreateRowHeights(table, table.HeightPoints);
        Dictionary<(int Row, int Column), PdfCore.PdfColor>? cellFills = CreateTableCellFills(table, includePresentationTableStyle);
        if (cellFills != null) {
            style.CellFills = MergeTableCellMap(style.CellFills, cellFills);
        }

        Dictionary<(int Row, int Column), PdfCore.PdfCellPadding>? cellPaddings = CreateTableCellPaddings(table);
        if (cellPaddings != null) {
            style.CellPaddings = MergeTableCellMap(style.CellPaddings, cellPaddings);
        }

        Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>? cellAlignments = CreateTableCellAlignments(table);
        if (cellAlignments != null) {
            style.CellAlignments = MergeTableCellMap(style.CellAlignments, cellAlignments);
        }

        Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>? cellVerticalAlignments = CreateTableCellVerticalAlignments(table);
        if (cellVerticalAlignments != null) {
            style.CellVerticalAlignments = MergeTableCellMap(style.CellVerticalAlignments, cellVerticalAlignments);
        }

        Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? cellBorders = CreateTableCellBorders(table, includePresentationTableStyle);
        if (cellBorders != null) {
            style.CellBorders = MergeTableCellMap(style.CellBorders, cellBorders);
        }

        return style;
    }

    private static PdfCore.PdfTableStyle CreateBaseTableStyle(PowerPointPdfExportOptions options) {
        PdfCore.PdfTableStyle? configuredStyle = options.PdfOptions?.HasExplicitDefaultTableStyle == true
            ? options.PdfOptions.DefaultTableStyle
            : null;
        if (configuredStyle != null) {
            return configuredStyle.Clone();
        }

        var style = new PdfCore.PdfTableStyle {
            HeaderFill = null,
            FooterFill = null,
            RowStripeFill = null,
            TextColor = PdfCore.PdfColor.Black,
            HeaderBold = false,
            FooterBold = false,
            LineHeight = 1.2D
        };
        // Every anchor cell receives its resolved presentation borders. A PDF
        // preset grid would paint over them and restore explicitly hidden edges.
        style.BorderColor = null;
        style.FontSize = PptCore.PowerPointTextDefaults.DefaultFontSizePoints;
        style.HeaderFontSize = PptCore.PowerPointTextDefaults.DefaultFontSizePoints;
        // DrawingML tcPr defaults: 0.1in on each side and 0.05in above/below.
        style.CellPaddingX = 7.2D;
        style.CellPaddingY = 3.6D;
        return style;
    }

    private static Dictionary<(int Row, int Column), T> MergeTableCellMap<T>(Dictionary<(int Row, int Column), T>? baseline, Dictionary<(int Row, int Column), T> overlay) {
        var merged = baseline == null
            ? new Dictionary<(int Row, int Column), T>()
            : new Dictionary<(int Row, int Column), T>(baseline);
        foreach (KeyValuePair<(int Row, int Column), T> item in overlay) {
            merged[item.Key] = item.Value;
        }

        return merged;
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

    private static Dictionary<(int Row, int Column), PdfCore.PdfColor>? CreateTableCellFills(PptCore.PowerPointTable table, bool includePresentationTableStyle) {
        var fills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            if (includePresentationTableStyle) {
                OfficeColor fill = PptCore.PowerPointSlideImageRenderer.ResolveTableCellFillColorForExport(table, row, column);
                if (fill.A > 0) {
                    fills[(row, column)] = PdfCore.PdfColor.FromOfficeColor(fill);
                }

                return;
            }

            PdfCore.PdfColor? directFill = ParsePdfColor(cell.FillColor);
            if (directFill.HasValue) {
                fills[(row, column)] = directFill.Value;
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
            PdfCore.PdfColumnAlign? align = MapColumnAlign(cell.HorizontalAlignment.ToOpenXml());
            if (align.HasValue) {
                alignments[(row, column)] = align.Value;
            }
        });

        return alignments.Count == 0 ? null : alignments;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>? CreateTableCellVerticalAlignments(PptCore.PowerPointTable table) {
        var alignments = new Dictionary<(int Row, int Column), PdfCore.PdfCellVerticalAlign>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            PdfCore.PdfCellVerticalAlign? align = MapVerticalAlign(cell.VerticalAlignment.ToOpenXml());
            if (align.HasValue) {
                alignments[(row, column)] = align.Value;
            }
        });

        return alignments.Count == 0 ? null : alignments;
    }

    private static Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>? CreateTableCellBorders(PptCore.PowerPointTable table, bool includePresentationTableStyle) {
        var borders = new Dictionary<(int Row, int Column), PdfCore.PdfCellBorder>();
        ForEachTableAnchorCell(table, (row, column, cell) => {
            if (includePresentationTableStyle) {
                OfficeBorderBox border = PptCore.PowerPointSlideImageRenderer.ResolveTableCellBordersForExport(table, row, column);
                borders[(row, column)] = CreatePdfTableCellBorder(border);
                return;
            }

            PdfCore.PdfColor? directBorder = ParsePdfColor(cell.BorderColor);
            if (directBorder.HasValue) {
                borders[(row, column)] = new PdfCore.PdfCellBorder {
                    Color = directBorder.Value,
                    Width = 0.75D
                };
            }
        });

        return borders.Count == 0 ? null : borders;
    }

    private static PdfCore.PdfCellBorder CreatePdfTableCellBorder(OfficeBorderBox border) {
        return new PdfCore.PdfCellBorder {
            Color = null,
            Width = 0D,
            LeftBorder = CreatePdfTableCellBorderSide(border.Left),
            TopBorder = CreatePdfTableCellBorderSide(border.Top),
            RightBorder = CreatePdfTableCellBorderSide(border.Right),
            BottomBorder = CreatePdfTableCellBorderSide(border.Bottom),
            DiagonalDownBorder = CreatePdfTableCellBorderSide(border.DiagonalDown),
            DiagonalUpBorder = CreatePdfTableCellBorderSide(border.DiagonalUp),
            DiagonalDown = border.DiagonalDown.HasValue,
            DiagonalUp = border.DiagonalUp.HasValue
        };
    }

    private static PdfCore.PdfCellBorderSide? CreatePdfTableCellBorderSide(OfficeBorderSide? side) {
        if (!side.HasValue) {
            return null;
        }

        OfficeBorderSide value = side.Value;
        return new PdfCore.PdfCellBorderSide {
            Color = value.IsVisible ? PdfCore.PdfColor.FromOfficeColor(value.Color) : null,
            Width = value.Width,
            DashStyle = value.DashStyle,
            LineStyle = value.LineKind == OfficeBorderLineKind.Double
                ? PdfCore.PdfCellBorderLineStyle.TwoLine
                : PdfCore.PdfCellBorderLineStyle.Standard
        };
    }

    private static void ForEachTableAnchorCell(PptCore.PowerPointTable table, Action<int, int, PptCore.PowerPointTableCell> action) {
        foreach (var anchor in table.EnumerateAnchorCells()) action(anchor.Row, anchor.Column, anchor.Cell);
    }

    private static (double Width, double Height) ResolveTableGridSize(PptCore.PowerPointTable table, double width, double height) {
        long frameWidth = table.Width, frameHeight = table.Height;
        long gridWidth = 0, gridHeight = 0;
        for (int c = 0; c < table.Columns; c++) gridWidth += table.GetColumnWidth(c);
        for (int r = 0; r < table.Rows; r++) gridHeight += table.GetRowHeight(r);
        double resolvedWidth = gridWidth > 0 && frameWidth > 0 ? width * gridWidth / frameWidth : width;
        double resolvedHeight = gridHeight > 0 && frameHeight > 0 ? height * gridHeight / frameHeight : height;
        return (resolvedWidth, resolvedHeight);
    }
}
