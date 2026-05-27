using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private const double TableCellClipBleed = 2D;

    // Helper shapes for column pagination
    private abstract class ColItem { public string Kind = string.Empty; }
    private sealed class ColPar : ColItem { public RichParagraphBlock Block = null!; public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines = null!; public System.Collections.Generic.List<double> Heights = null!; public double Leading; public double Size; public double XOffset; public double TextWidth; public double FirstLineXOffset; public double FirstLineTextWidth; public ColPar() { Kind = "P"; } }
    private sealed class ColHead : ColItem { public HeadingBlock Block = null!; public System.Collections.Generic.List<string> Lines = null!; public double Leading; public double Size; public double SpacingBefore; public double SpacingAfter; public bool KeepWithNext; public PdfColor? Color; public ColHead() { Kind = "H"; } }
    private sealed class ColRule : ColItem { public HorizontalRuleBlock Block = null!; public ColRule() { Kind = "R"; } }
    private sealed class ColImg : ColItem { public ImageBlock Block = null!; public ColImg() { Kind = "I"; } }
    private sealed class ColShape : ColItem { public ShapeBlock Block = null!; public ColShape() { Kind = "S"; } }
    private sealed class ColDrawing : ColItem { public DrawingBlock Block = null!; public ColDrawing() { Kind = "D"; } }
    private sealed class ColBookmark : ColItem { public BookmarkBlock Block = null!; public ColBookmark() { Kind = "B"; } }
    private sealed class ColSpacer : ColItem { public SpacerBlock Block = null!; public ColSpacer() { Kind = "SPACE"; } }
    private sealed class ColListItem : ColItem { public System.Collections.Generic.List<string> Lines = null!; public string Marker = string.Empty; public double MarkerXOffset; public double MarkerWidth; public PdfAlign MarkerAlign; public double TextXOffset; public double TextWidth; public PdfAlign TextAlign; public PdfColor? Color; public double Leading; public double Size; public double SpacingBefore; public double SpacingAfter; public bool KeepTogether; public bool IsFirstInKeepGroup; public double KeepGroupHeight; public bool KeepWithNext; public bool IsFirstInKeepWithNextGroup; public int KeepWithNextGroupItemCount; public double KeepWithNextGroupHeight; public ColListItem() { Kind = "L"; } }
    private sealed class ColPanel : ColItem { public PanelParagraphBlock Block = null!; public PanelStyle Style = null!; public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines = null!; public System.Collections.Generic.List<double> Heights = null!; public double Leading; public double Size; public double FirstBaselineOffset; public double XOffset; public double PanelWidth; public double TextWidth; public ColPanel() { Kind = "PANEL"; } }
    private sealed class ColTable : ColItem { public TableBlock Block = null!; public PdfTableStyle Style = null!; public int Columns; public double[] ColumnWidths = null!; public string[][][] RowLines = null!; public int[] RowLineCounts = null!; public double[] RowHeights = null!; public double[] RowLeadings = null!; public double[] RowSizes = null!; public bool[] RowBold = null!; public double Width; public double Size; public int HeaderRowCount; public int FooterStartRowIndex; public System.Collections.Generic.List<string>? CaptionLines; public double CaptionLeading; public double CaptionHeight; public ColTable() { Kind = "T"; } }
    private sealed class TableColumnLayout { public double[] Widths = null!; public double Width; }
    private readonly struct TableCellLayout {
        public TableCellLayout(int column, int columnSpan, int rowSpan, string text, string? linkUri, string? linkContents) {
            Column = column;
            ColumnSpan = columnSpan;
            RowSpan = rowSpan;
            Text = text;
            LinkUri = linkUri;
            LinkContents = linkContents;
        }

        public int Column { get; }
        public int ColumnSpan { get; }
        public int RowSpan { get; }
        public string Text { get; }
        public string? LinkUri { get; }
        public string? LinkContents { get; }
    }
    private static double GetParagraphLeading(PdfParagraphStyle? style, double fontSize) {
        double multiplier = style?.LineHeight ?? 1.4;
        if (multiplier <= 0 || double.IsNaN(multiplier) || double.IsInfinity(multiplier)) {
            throw new ArgumentException("Paragraph line height must be a positive finite value.");
        }

        return fontSize * multiplier;
    }

    private static double GetParagraphSpacingBefore(PdfParagraphStyle? style) {
        double spacingBefore = style?.SpacingBefore ?? 0;
        if (spacingBefore < 0 || double.IsNaN(spacingBefore) || double.IsInfinity(spacingBefore)) {
            throw new ArgumentException("Paragraph spacing before must be a non-negative finite value.");
        }

        return spacingBefore;
    }

    private static double GetParagraphSpacingAfter(PdfParagraphStyle? style, double leading) {
        double spacingAfter = style?.SpacingAfter ?? leading * 0.3;
        if (spacingAfter < 0 || double.IsNaN(spacingAfter) || double.IsInfinity(spacingAfter)) {
            throw new ArgumentException("Paragraph spacing after must be a non-negative finite value.");
        }

        return spacingAfter;
    }

    private static double GetParagraphTabStopWidth(PdfParagraphStyle? style) {
        double tabStopWidth = style?.DefaultTabStopWidth ?? DefaultParagraphTabStopWidth;
        if (tabStopWidth <= 0 || double.IsNaN(tabStopWidth) || double.IsInfinity(tabStopWidth)) {
            throw new ArgumentException("Paragraph default tab stop width must be a positive finite value.");
        }

        return tabStopWidth;
    }

    private static PdfHeadingStyle? ResolveHeadingStyle(HeadingBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultHeadingStylesSnapshot?.GetSnapshot(block.Level);
    }

    private static double GetHeadingFontSize(HeadingBlock block, PdfHeadingStyle? style) {
        return style?.GetFontSize(block.Level) ?? PdfHeadingStyle.GetDefaultFontSize(block.Level);
    }

    private static double GetHeadingLeading(PdfHeadingStyle? style, double fontSize) {
        return style?.GetLeading(fontSize) ?? fontSize * 1.25D;
    }

    private static double GetHeadingSpacingAfter(PdfHeadingStyle? style, double leading) {
        return style?.GetSpacingAfter(leading) ?? leading * 0.25D;
    }

    private static PdfListStyle? ResolveListStyle(BulletListBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultListStyleSnapshot;
    }

    private static PdfListStyle? ResolveListStyle(NumberedListBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultListStyleSnapshot;
    }

    private static double GetListFontSize(PdfListStyle? style, double defaultFontSize) {
        return style?.GetFontSize(defaultFontSize) ?? defaultFontSize;
    }

    private static double GetListLeading(PdfListStyle? style, double fontSize) {
        return style?.GetLeading(fontSize) ?? fontSize * 1.4D;
    }

    private static double GetListMarkerGap(PdfListStyle? style, double defaultGap) {
        return style?.GetMarkerGap(defaultGap) ?? defaultGap;
    }

    private static double GetListItemSpacing(PdfListStyle? style, double leading) {
        return style?.GetItemSpacing(leading) ?? leading * 0.15D;
    }

    private static double GetTableBodyFontSize(PdfTableStyle style, double defaultFontSize) {
        return style.FontSize ?? defaultFontSize;
    }

    private static double GetTableLeading(PdfTableStyle style, double fontSize) {
        double multiplier = style.LineHeight ?? 1.4D;
        if (multiplier <= 0 || double.IsNaN(multiplier) || double.IsInfinity(multiplier)) {
            throw new ArgumentException("Table line height must be a positive finite value.");
        }

        return fontSize * multiplier;
    }

    private static double GetTableCellPaddingLeft(PdfTableStyle style) {
        return style.CellPaddingLeft ?? style.CellPaddingX;
    }

    private static double GetTableCellPaddingRight(PdfTableStyle style) {
        return style.CellPaddingRight ?? style.CellPaddingX;
    }

    private static double GetTableCellPaddingTop(PdfTableStyle style) {
        return style.CellPaddingTop ?? style.CellPaddingY;
    }

    private static double GetTableCellPaddingBottom(PdfTableStyle style) {
        return style.CellPaddingBottom ?? style.CellPaddingY;
    }

    private static int GetTableColumnCount(TableBlock table) => table.ColumnCount;

    private static void ValidateTableRoleRowCounts(PdfTableStyle style, int rowCount) {
        if (style.HeaderRowCount > rowCount) {
            throw new ArgumentException("Table header row count cannot exceed the table row count.");
        }

        if (style.FooterRowCount > rowCount) {
            throw new ArgumentException("Table footer row count cannot exceed the table row count.");
        }

        if (style.FooterRowCount > rowCount - style.HeaderRowCount) {
            throw new ArgumentException("Table header and footer row counts cannot exceed the table row count.");
        }
    }

    private static void ValidateTableCellStyleCoordinates(PdfTableStyle style, int rowCount, int columnCount) {
        if (style.CellFills != null) {
            foreach (var cellFill in style.CellFills) {
                if (cellFill.Key.Row < 0 || cellFill.Key.Column < 0) {
                    throw new ArgumentException("Table cell fill coordinates cannot be negative.");
                }

                if (cellFill.Key.Row >= rowCount || cellFill.Key.Column >= columnCount) {
                    throw new ArgumentException("Table cell fill coordinates must fit inside the table grid.");
                }
            }
        }

        if (style.CellBorders != null) {
            foreach (var cellBorder in style.CellBorders) {
                if (cellBorder.Key.Row < 0 || cellBorder.Key.Column < 0) {
                    throw new ArgumentException("Table cell border coordinates cannot be negative.");
                }

                if (cellBorder.Key.Row >= rowCount || cellBorder.Key.Column >= columnCount) {
                    throw new ArgumentException("Table cell border coordinates must fit inside the table grid.");
                }
            }
        }
    }

    private static void ValidateTableColumnStyleBounds(PdfTableStyle style, int columnCount) {
        if (style.BodyColumnFills != null) {
            for (int column = columnCount; column < style.BodyColumnFills.Count; column++) {
                if (style.BodyColumnFills[column] != null) {
                    throw new ArgumentException("Table body column fills must fit inside the table grid.");
                }
            }
        }

        if (style.Alignments != null && style.Alignments.Count > columnCount) {
            throw new ArgumentException("Table column alignments must fit inside the table grid.");
        }

        if (style.VerticalAlignments != null && style.VerticalAlignments.Count > columnCount) {
            throw new ArgumentException("Table vertical alignments must fit inside the table grid.");
        }

        ValidateOptionalColumnStyleBounds(style.ColumnWidthPoints, columnCount, "Table fixed column widths must fit inside the table grid.");
        ValidateOptionalColumnStyleBounds(style.ColumnMinWidthPoints, columnCount, "Table minimum column widths must fit inside the table grid.");
        ValidateOptionalColumnStyleBounds(style.ColumnMaxWidthPoints, columnCount, "Table maximum column widths must fit inside the table grid.");

        if (style.ColumnWidthWeights != null && style.ColumnWidthWeights.Count > columnCount) {
            throw new ArgumentException("Table column width weights must fit inside the table grid.");
        }
    }

    private static void ValidateOptionalColumnStyleBounds(System.Collections.Generic.List<double?>? values, int columnCount, string message) {
        if (values == null) {
            return;
        }

        for (int column = columnCount; column < values.Count; column++) {
            if (values[column].HasValue) {
                throw new ArgumentException(message);
            }
        }
    }

    private static System.Collections.Generic.List<TableCellLayout> GetTableCellLayouts(TableBlock table, int rowIndex, int columnCount) {
        var targetCells = new System.Collections.Generic.List<TableCellLayout>();
        if (rowIndex < 0 || rowIndex >= table.Cells.Count) {
            return targetCells;
        }

        var activeRowSpans = new int[columnCount];
        for (int currentRow = 0; currentRow <= rowIndex; currentRow++) {
            int column = 0;
            var row = table.Cells[currentRow];
            for (int cellIndex = 0; cellIndex < row.Count && column < columnCount; cellIndex++) {
                while (column < columnCount && activeRowSpans[column] > 0) {
                    column++;
                }

                if (column >= columnCount) {
                    break;
                }

                PdfTableCell cell = row[cellIndex];
                int columnSpan = System.Math.Min(cell.ColumnSpan, columnCount - column);
                int rowSpan = System.Math.Min(cell.RowSpan, table.Cells.Count - currentRow);
                if (currentRow == rowIndex) {
                    targetCells.Add(new TableCellLayout(column, columnSpan, rowSpan, cell.Text, cell.LinkUri, cell.LinkContents));
                }

                for (int c = column; c < column + columnSpan; c++) {
                    activeRowSpans[c] = System.Math.Max(activeRowSpans[c], rowSpan);
                }

                column += columnSpan;
            }

            for (int c = 0; c < activeRowSpans.Length; c++) {
                if (activeRowSpans[c] > 0) {
                    activeRowSpans[c]--;
                }
            }
        }

        return targetCells;
    }

    private static double GetTableCellWidth(double[] columnWidths, int column, int columnSpan, double columnGap) {
        double width = 0D;
        int lastColumn = System.Math.Min(columnWidths.Length, column + columnSpan);
        for (int index = column; index < lastColumn; index++) {
            width += columnWidths[index];
            if (index > column) {
                width += columnGap;
            }
        }

        return width;
    }

    private static double GetTableCellHeight(double[] rowHeights, int row, int rowSpan) {
        double height = 0D;
        int lastRow = System.Math.Min(rowHeights.Length, row + rowSpan);
        for (int index = row; index < lastRow; index++) {
            height += rowHeights[index];
        }

        return height;
    }

    private static void ApplyTableRowSpanHeights(TableBlock table, int columnCount, string[][][] rowLines, double[] rowHeights, double[] rowLeadings, double padTop, double padBottom) {
        for (int rowIndex = 0; rowIndex < table.Cells.Count; rowIndex++) {
            var cells = GetTableCellLayouts(table, rowIndex, columnCount);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                if (cell.RowSpan <= 1) {
                    continue;
                }

                int rowSpan = System.Math.Min(cell.RowSpan, rowHeights.Length - rowIndex);
                if (rowSpan <= 1) {
                    continue;
                }

                string[] lines = rowLines[rowIndex][cell.Column];
                double requiredHeight = System.Math.Max(1, lines.Length) * rowLeadings[rowIndex] + padTop + padBottom;
                double currentHeight = GetTableCellHeight(rowHeights, rowIndex, rowSpan);
                if (requiredHeight <= currentHeight + 0.001) {
                    continue;
                }

                double extraPerRow = (requiredHeight - currentHeight) / rowSpan;
                for (int spanRow = rowIndex; spanRow < rowIndex + rowSpan; spanRow++) {
                    rowHeights[spanRow] += extraPerRow;
                }
            }
        }
    }

    private static void ValidateTableRowSpansWithinRoleBoundaries(TableBlock table, int columnCount, int headerRowCount, int footerStartRowIndex) {
        for (int rowIndex = 0; rowIndex < table.Cells.Count; rowIndex++) {
            var cells = GetTableCellLayouts(table, rowIndex, columnCount);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                if (cell.RowSpan <= 1) {
                    continue;
                }

                int lastRowExclusive = rowIndex + cell.RowSpan;
                if (rowIndex < headerRowCount && lastRowExclusive > headerRowCount) {
                    throw new ArgumentException("Table cell row span cannot cross the table header boundary.");
                }

                if (rowIndex < footerStartRowIndex && lastRowExclusive > footerStartRowIndex) {
                    throw new ArgumentException("Table cell row span cannot cross the table footer boundary.");
                }
            }
        }
    }

    private static bool TryGetTableCellLayoutAtColumn(System.Collections.Generic.List<TableCellLayout> cells, int column, out TableCellLayout layout) {
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            if (cells[cellIndex].Column == column) {
                layout = cells[cellIndex];
                return true;
            }
        }

        layout = default;
        return false;
    }

    private static bool IsTableBoundaryInsideSpannedCell(TableBlock table, int rowIndex, int boundaryColumn, int columnCount) {
        if (rowIndex < 0 || rowIndex >= table.Cells.Count || boundaryColumn < 0 || boundaryColumn >= columnCount - 1) {
            return false;
        }

        for (int sourceRowIndex = 0; sourceRowIndex <= rowIndex; sourceRowIndex++) {
            var cells = GetTableCellLayouts(table, sourceRowIndex, columnCount);
            for (int i = 0; i < cells.Count; i++) {
                TableCellLayout cell = cells[i];
                if (sourceRowIndex + cell.RowSpan <= rowIndex) {
                    continue;
                }

                if (cell.Column <= boundaryColumn && boundaryColumn < cell.Column + cell.ColumnSpan - 1) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool[] GetRowSpanBoundarySkipColumns(TableBlock table, int boundaryRowIndex, int columnCount) {
        var skipColumns = new bool[columnCount];
        if (boundaryRowIndex < 0 || boundaryRowIndex >= table.Cells.Count - 1 || columnCount <= 0) {
            return skipColumns;
        }

        for (int rowIndex = 0; rowIndex <= boundaryRowIndex; rowIndex++) {
            var cells = GetTableCellLayouts(table, rowIndex, columnCount);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                if (cell.RowSpan <= 1 || boundaryRowIndex >= rowIndex + cell.RowSpan - 1) {
                    continue;
                }

                int lastColumn = System.Math.Min(columnCount, cell.Column + cell.ColumnSpan);
                for (int column = cell.Column; column < lastColumn; column++) {
                    skipColumns[column] = true;
                }
            }
        }

        return skipColumns;
    }

    private static bool[] GetRowSpanContinuationSkipColumns(TableBlock table, int rowIndex, int columnCount) {
        var skipColumns = new bool[columnCount];
        if (rowIndex <= 0 || rowIndex >= table.Cells.Count || columnCount <= 0) {
            return skipColumns;
        }

        for (int startRow = 0; startRow < rowIndex; startRow++) {
            var cells = GetTableCellLayouts(table, startRow, columnCount);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                if (cell.RowSpan <= rowIndex - startRow) {
                    continue;
                }

                int lastColumn = System.Math.Min(columnCount, cell.Column + cell.ColumnSpan);
                for (int column = cell.Column; column < lastColumn; column++) {
                    skipColumns[column] = true;
                }
            }
        }

        return skipColumns;
    }

    private static bool[] GetMergedCellContinuationSkipColumns(TableBlock table, int rowIndex, int columnCount) {
        bool[] skipColumns = GetRowSpanContinuationSkipColumns(table, rowIndex, columnCount);
        if (rowIndex < 0 || rowIndex >= table.Cells.Count || columnCount <= 0) {
            return skipColumns;
        }

        var cells = GetTableCellLayouts(table, rowIndex, columnCount);
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            TableCellLayout cell = cells[cellIndex];
            int lastColumn = System.Math.Min(columnCount, cell.Column + cell.ColumnSpan);
            for (int column = cell.Column + 1; column < lastColumn; column++) {
                skipColumns[column] = true;
            }
        }

        return skipColumns;
    }

    private static void DrawTableHorizontalLine(StringBuilder sb, PdfColor color, double width, double xOrigin, double[] columnWidths, double columnGap, double y, bool[] skipColumns) {
        if (columnWidths.Length == 0) {
            return;
        }

        double tableWidth = GetTableCellWidth(columnWidths, 0, columnWidths.Length, columnGap);
        if (!HasSkippedColumns(skipColumns, columnWidths.Length)) {
            DrawHLine(sb, color, width, xOrigin, xOrigin + tableWidth, y);
            return;
        }

        var columnLefts = new double[columnWidths.Length];
        var columnRights = new double[columnWidths.Length];
        double x = xOrigin;
        for (int column = 0; column < columnWidths.Length; column++) {
            columnLefts[column] = x;
            columnRights[column] = x + columnWidths[column];
            x += columnWidths[column] + columnGap;
        }

        int segmentStart = -1;
        for (int column = 0; column < columnWidths.Length; column++) {
            bool skip = column < skipColumns.Length && skipColumns[column];
            if (skip) {
                if (segmentStart >= 0) {
                    DrawHLine(sb, color, width, columnLefts[segmentStart], columnRights[column - 1], y);
                    segmentStart = -1;
                }

                continue;
            }

            if (segmentStart < 0) {
                segmentStart = column;
            }
        }

        if (segmentStart >= 0) {
            DrawHLine(sb, color, width, columnLefts[segmentStart], columnRights[columnWidths.Length - 1], y);
        }
    }

    private static void DrawTableRowFill(StringBuilder sb, PdfColor color, double xOrigin, double[] columnWidths, double columnGap, double y, double height, bool[] skipColumns) {
        if (columnWidths.Length == 0) {
            return;
        }

        double tableWidth = GetTableCellWidth(columnWidths, 0, columnWidths.Length, columnGap);
        if (!HasSkippedColumns(skipColumns, columnWidths.Length)) {
            DrawRowFill(sb, color, xOrigin, y, tableWidth, height);
            return;
        }

        var columnLefts = new double[columnWidths.Length];
        var columnRights = new double[columnWidths.Length];
        double x = xOrigin;
        for (int column = 0; column < columnWidths.Length; column++) {
            columnLefts[column] = x;
            columnRights[column] = x + columnWidths[column];
            x += columnWidths[column] + columnGap;
        }

        int segmentStart = -1;
        for (int column = 0; column < columnWidths.Length; column++) {
            bool skip = column < skipColumns.Length && skipColumns[column];
            if (skip) {
                if (segmentStart >= 0) {
                    DrawRowFill(sb, color, columnLefts[segmentStart], y, columnRights[column - 1] - columnLefts[segmentStart], height);
                    segmentStart = -1;
                }

                continue;
            }

            if (segmentStart < 0) {
                segmentStart = column;
            }
        }

        if (segmentStart >= 0) {
            DrawRowFill(sb, color, columnLefts[segmentStart], y, columnRights[columnWidths.Length - 1] - columnLefts[segmentStart], height);
        }
    }

    private static bool HasSkippedColumns(bool[] skipColumns, int columnCount) {
        for (int column = 0; column < columnCount && column < skipColumns.Length; column++) {
            if (skipColumns[column]) {
                return true;
            }
        }

        return false;
    }

    private static bool ShouldClipTableCellText(double textX, double textBaselineY, double textWidth, PdfStandardFont font, double fontSize, double cellX, double cellY, double cellWidth, double cellHeight) {
        const double epsilon = 0.01D;
        double ascender = GetAscender(font, fontSize);
        double descender = GetDescender(font, fontSize);

        return textX < cellX - epsilon ||
               textX + textWidth > cellX + cellWidth + epsilon ||
               textBaselineY + ascender > cellY + cellHeight + epsilon ||
               textBaselineY + descender < cellY - epsilon;
    }

    private static double GetTableRowFontSize(PdfTableStyle style, int rowIndex, int headerRowCount, int footerStartRowIndex, double defaultFontSize) {
        if (rowIndex < headerRowCount) {
            return style.HeaderFontSize ?? GetTableBodyFontSize(style, defaultFontSize);
        }

        if (rowIndex >= footerStartRowIndex) {
            return style.FooterFontSize ?? GetTableBodyFontSize(style, defaultFontSize);
        }

        return GetTableBodyFontSize(style, defaultFontSize);
    }

    private static bool GetTableRowBold(PdfTableStyle style, int rowIndex, int headerRowCount, int footerStartRowIndex) {
        return rowIndex < headerRowCount ? style.HeaderBold : rowIndex >= footerStartRowIndex && style.FooterBold;
    }

    private static PdfStandardFont GetTableRowFont(PdfOptions options, bool bold) {
        var normalFont = ChooseNormal(options.DefaultFont);
        return bold ? ChooseBold(normalFont) : normalFont;
    }

    private static string GetTableRowFontResource(bool bold) {
        return bold ? "F2" : "F1";
    }

    private static bool TableUsesBold(PdfTableStyle style, int rowCount, int headerRowCount, int footerStartRowIndex) {
        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            if (GetTableRowBold(style, rowIndex, headerRowCount, footerStartRowIndex)) {
                return true;
            }
        }

        return false;
    }

    private static PanelStyle ResolvePanelStyle(PanelParagraphBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultPanelStyleSnapshot ?? new PanelStyle();
    }

    private static PdfHorizontalRuleStyle ResolveHorizontalRuleStyle(HorizontalRuleBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultHorizontalRuleStyleSnapshot ?? new PdfHorizontalRuleStyle();
    }

    private static PdfImageStyle ResolveImageStyle(ImageBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultImageStyleSnapshot ?? new PdfImageStyle();
    }

    private static PdfDrawingStyle ResolveDrawingStyle(ShapeBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultDrawingStyleSnapshot ?? new PdfDrawingStyle();
    }

    private static PdfDrawingStyle ResolveDrawingStyle(DrawingBlock block, PdfOptions options) {
        return block.Style ?? options.DefaultDrawingStyleSnapshot ?? new PdfDrawingStyle();
    }

    private static (double X, double Width, double FirstLineX, double FirstLineWidth) GetParagraphTextFrame(PdfParagraphStyle? style, double x, double width) {
        double leftIndent = style?.LeftIndent ?? 0;
        double rightIndent = style?.RightIndent ?? 0;
        double firstLineIndent = style?.FirstLineIndent ?? 0;
        if (leftIndent < 0 || double.IsNaN(leftIndent) || double.IsInfinity(leftIndent)) {
            throw new ArgumentException("Paragraph left indent must be a non-negative finite value.");
        }

        if (rightIndent < 0 || double.IsNaN(rightIndent) || double.IsInfinity(rightIndent)) {
            throw new ArgumentException("Paragraph right indent must be a non-negative finite value.");
        }

        if (double.IsNaN(firstLineIndent) || double.IsInfinity(firstLineIndent)) {
            throw new ArgumentException("Paragraph first line indent must be a finite value.");
        }

        if (leftIndent + firstLineIndent < 0) {
            throw new ArgumentException("Paragraph first line indent must not move text outside the left content frame.");
        }

        double textWidth = width - leftIndent - rightIndent;
        if (textWidth <= 0 || double.IsNaN(textWidth) || double.IsInfinity(textWidth)) {
            throw new ArgumentException("Paragraph left and right indents must leave a positive text width.");
        }

        double firstLineWidth = textWidth - firstLineIndent;
        if (firstLineWidth <= 0 || double.IsNaN(firstLineWidth) || double.IsInfinity(firstLineWidth)) {
            throw new ArgumentException("Paragraph first line indent must leave a positive text width.");
        }

        return (x + leftIndent, textWidth, x + leftIndent + firstLineIndent, firstLineWidth);
    }

    private static bool TryApplyWidowControl(PdfParagraphStyle? style, int totalLineCount, int startLineIndex, ref int take, ref double heightSum, System.Collections.Generic.List<double> lineHeights, bool canMoveToNextPage) {
        if (style?.WidowControl != true || take <= 0) {
            return false;
        }

        int remainingLineCount = totalLineCount - startLineIndex;
        int afterTake = remainingLineCount - take;
        if (afterTake <= 0) {
            return false;
        }

        if (take == 1 && canMoveToNextPage) {
            return true;
        }

        if (afterTake == 1) {
            if (take > 2) {
                take--;
                heightSum -= lineHeights[startLineIndex + take];
            } else if (canMoveToNextPage) {
                return true;
            }
        }

        return false;
    }

    private static string BuildFooter(PdfOptions opts, int variantPage, int page, int pages, PdfStandardFont footerFont, string footerFontResource) {
        string text;
        var footerSegments = opts.GetFooterSegmentsForPage(variantPage);
        var footerZones = opts.GetFooterZonesForPage(variantPage);
        if (HasPageTextZones(footerZones)) {
            return BuildPageTextZones(opts, footerZones, page, pages, footerFont, footerFontResource, opts.FooterFontSize, opts.FooterTextColor, opts.FooterOffsetY, isHeader: false);
        } else if (footerSegments != null && footerSegments.Count > 0) {
            text = BuildPageTextFromSegments(footerSegments, page, pages, opts.PageNumberStyle);
        } else {
            text = FormatPageText(opts.GetFooterFormatForPage(variantPage), page, pages, opts.PageNumberStyle);
        }
        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double textWidth = EstimateSimpleTextWidth(text, footerFont, opts.FooterFontSize);
        double x = opts.MarginLeft;
        if (opts.FooterAlign == PdfAlign.Center) x = opts.MarginLeft + Math.Max(0, (width - textWidth) / 2);
        else if (opts.FooterAlign == PdfAlign.Right) x = opts.MarginLeft + Math.Max(0, width - textWidth);
        double y = opts.MarginBottom - opts.FooterOffsetY;
        PdfColor? footerColor = opts.FooterTextColor;
        var sb = new StringBuilder();
        var content = new ContentStreamBuilder(sb)
            .BeginText()
            .Font(footerFontResource, opts.FooterFontSize);
        if (footerColor.HasValue) {
            content.FillColor(footerColor.Value);
        }

        content
            .TextMatrix(x, y)
            .ShowHexText(EncodeWinAnsiHex(text))
            .EndText();
        return sb.ToString();
    }

    private static string BuildHeader(PdfOptions opts, int variantPage, int page, int pages, PdfStandardFont headerFont, string headerFontResource) {
        string text;
        var headerSegments = opts.GetHeaderSegmentsForPage(variantPage);
        var headerZones = opts.GetHeaderZonesForPage(variantPage);
        if (HasPageTextZones(headerZones)) {
            return BuildPageTextZones(opts, headerZones, page, pages, headerFont, headerFontResource, opts.HeaderFontSize, opts.HeaderTextColor, opts.HeaderOffsetY, isHeader: true);
        } else if (headerSegments != null && headerSegments.Count > 0) {
            text = BuildPageTextFromSegments(headerSegments, page, pages, opts.PageNumberStyle);
        } else {
            text = FormatPageText(opts.GetHeaderFormatForPage(variantPage), page, pages, opts.PageNumberStyle);
        }

        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double textWidth = EstimateSimpleTextWidth(text, headerFont, opts.HeaderFontSize);
        double x = opts.MarginLeft;
        if (opts.HeaderAlign == PdfAlign.Center) x = opts.MarginLeft + Math.Max(0, (width - textWidth) / 2);
        else if (opts.HeaderAlign == PdfAlign.Right) x = opts.MarginLeft + Math.Max(0, width - textWidth);
        double y = opts.PageHeight - opts.MarginTop + opts.HeaderOffsetY;
        PdfColor? headerColor = opts.HeaderTextColor;

        var sb = new StringBuilder();
        var content = new ContentStreamBuilder(sb)
            .BeginText()
            .Font(headerFontResource, opts.HeaderFontSize);
        if (headerColor.HasValue) {
            content.FillColor(headerColor.Value);
        }

        content
            .TextMatrix(x, y)
            .ShowHexText(EncodeWinAnsiHex(text))
            .EndText();
        return sb.ToString();
    }

    private static bool HasPageTextZones((string? Left, string? Center, string? Right) zones) =>
        !string.IsNullOrEmpty(zones.Left) ||
        !string.IsNullOrEmpty(zones.Center) ||
        !string.IsNullOrEmpty(zones.Right);

    private static string BuildPageTextZones(
        PdfOptions opts,
        (string? Left, string? Center, string? Right) zones,
        int page,
        int pages,
        PdfStandardFont font,
        string fontResource,
        double fontSize,
        PdfColor? color,
        double offset,
        bool isHeader) {
        double width = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        double y = isHeader ? opts.PageHeight - opts.MarginTop + offset : opts.MarginBottom - offset;
        var sb = new StringBuilder();
        var zoneLayouts = BuildPageTextZoneLayouts(opts, zones, page, pages, font, fontSize, isHeader);
        foreach (var zone in zoneLayouts) {
            AppendPageText(sb, zone.Text, fontResource, fontSize, color, zone.X, y);
        }

        return sb.ToString();
    }

    private static System.Collections.Generic.List<(string Name, string Text, double X, double Width)> BuildPageTextZoneLayouts(
        PdfOptions opts,
        (string? Left, string? Center, string? Right) zones,
        int page,
        int pages,
        PdfStandardFont font,
        double fontSize,
        bool isHeader) {
        double contentLeft = opts.MarginLeft;
        double contentWidth = opts.PageWidth - opts.MarginLeft - opts.MarginRight;
        var layouts = new System.Collections.Generic.List<(string Name, string Text, double X, double Width)>();

        if (!string.IsNullOrEmpty(zones.Left)) {
            string text = FormatPageText(zones.Left!, page, pages, opts.PageNumberStyle);
            double textWidth = EstimateSimpleTextWidth(text, font, fontSize);
            layouts.Add(("left", text, contentLeft, textWidth));
        }

        if (!string.IsNullOrEmpty(zones.Center)) {
            string text = FormatPageText(zones.Center!, page, pages, opts.PageNumberStyle);
            double textWidth = EstimateSimpleTextWidth(text, font, fontSize);
            layouts.Add(("center", text, contentLeft + ((contentWidth - textWidth) / 2), textWidth));
        }

        if (!string.IsNullOrEmpty(zones.Right)) {
            string text = FormatPageText(zones.Right!, page, pages, opts.PageNumberStyle);
            double textWidth = EstimateSimpleTextWidth(text, font, fontSize);
            layouts.Add(("right", text, contentLeft + contentWidth - textWidth, textWidth));
        }

        ValidatePageTextZoneLayouts(layouts, contentLeft, contentLeft + contentWidth, isHeader);
        return layouts;
    }

    private static void ValidatePageTextZoneLayouts(System.Collections.Generic.List<(string Name, string Text, double X, double Width)> layouts, double contentLeft, double contentRight, bool isHeader) {
        const double tolerance = 0.01D;
        const double minimumGap = 2D;
        string scope = isHeader ? "header" : "footer";
        foreach (var zone in layouts) {
            if (zone.X < contentLeft - tolerance || zone.X + zone.Width > contentRight + tolerance) {
                throw new ArgumentException("PDF " + scope + " zone text must fit inside the page content width.");
            }
        }

        var ordered = layouts.OrderBy(zone => zone.X).ToList();
        for (int i = 1; i < ordered.Count; i++) {
            var previous = ordered[i - 1];
            var current = ordered[i];
            if (previous.X + previous.Width + minimumGap > current.X + tolerance) {
                throw new ArgumentException("PDF " + scope + " zone text must not overlap.");
            }
        }
    }

    private static void AppendPageText(StringBuilder sb, string text, string fontResource, double fontSize, PdfColor? color, double x, double y) {
        var content = new ContentStreamBuilder(sb)
            .BeginText()
            .Font(fontResource, fontSize);
        if (color.HasValue) {
            content.FillColor(color.Value);
        }

        content
            .TextMatrix(x, y)
            .ShowHexText(EncodeWinAnsiHex(text))
            .EndText();
    }

    private static string BuildPageTextFromSegments(System.Collections.Generic.IReadOnlyList<FooterSegment> segments, int page, int pages, PdfPageNumberStyle style) {
        var sb = new StringBuilder();
        foreach (var segment in segments) {
            switch (segment.Kind) {
                case FooterSegmentKind.Text:
                    sb.Append(segment.Text);
                    break;
                case FooterSegmentKind.PageNumber:
                    sb.Append(FormatPageNumber(page, style));
                    break;
                case FooterSegmentKind.TotalPages:
                    sb.Append(FormatPageNumber(pages, style));
                    break;
            }
        }

        return sb.ToString();
    }

    private static string FormatPageText(string format, int page, int pages, PdfPageNumberStyle style) {
        string pageText = FormatPageNumber(page, style);
        string pagesText = FormatPageNumber(pages, style);
        return format
            .Replace("{page}", pageText)
            .Replace("{pages}", pagesText);
    }

    private static string FormatPageNumber(int number, PdfPageNumberStyle style) {
        Guard.PageNumberStyle(style, nameof(style));
        if (number < 1) {
            throw new ArgumentOutOfRangeException(nameof(number), "PDF page number must be positive.");
        }

        switch (style) {
            case PdfPageNumberStyle.Arabic:
                return number.ToString(CultureInfo.InvariantCulture);
            case PdfPageNumberStyle.LowerRoman:
                return ToRoman(number).ToLowerInvariant();
            case PdfPageNumberStyle.UpperRoman:
                return ToRoman(number);
            case PdfPageNumberStyle.LowerLetter:
                return ToLetters(number, upper: false);
            case PdfPageNumberStyle.UpperLetter:
                return ToLetters(number, upper: true);
            default:
                throw new ArgumentException("PDF page number style must be Arabic, LowerRoman, UpperRoman, LowerLetter, or UpperLetter.", nameof(style));
        }
    }

    private static string ToRoman(int number) {
        var values = new[] { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        var numerals = new[] { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        var sb = new StringBuilder();
        int remaining = number;
        for (int i = 0; i < values.Length; i++) {
            while (remaining >= values[i]) {
                sb.Append(numerals[i]);
                remaining -= values[i];
            }
        }

        return sb.ToString();
    }

    private static string ToLetters(int number, bool upper) {
        var chars = new System.Collections.Generic.List<char>();
        int remaining = number;
        char baseChar = upper ? 'A' : 'a';
        while (remaining > 0) {
            remaining--;
            chars.Add((char)(baseChar + (remaining % 26)));
            remaining /= 26;
        }

        chars.Reverse();
        return new string(chars.ToArray());
    }

    private static double? GetOptionalColumnWidth(System.Collections.Generic.List<double?>? values, int index, string errorMessage) {
        if (values == null || index >= values.Count || !values[index].HasValue) {
            return null;
        }

        double value = values[index]!.Value;
        if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentException(errorMessage);
        }

        return value;
    }

    private static double ResolveTableFrameWidth(PdfTableStyle style, double containerWidth) {
        if (style.LeftIndent < 0 || double.IsNaN(style.LeftIndent) || double.IsInfinity(style.LeftIndent)) {
            throw new ArgumentException("Table left indent must be a non-negative finite value.");
        }

        double frameWidth = containerWidth - style.LeftIndent;
        if (frameWidth <= 0.001 || double.IsNaN(frameWidth) || double.IsInfinity(frameWidth)) {
            throw new ArgumentException("Table left indent must leave a positive table width.");
        }

        return frameWidth;
    }

    private static double ResolveTableAvailableWidth(PdfTableStyle style, double containerWidth) {
        double frameWidth = ResolveTableFrameWidth(style, containerWidth);
        if (style.MaxWidth.HasValue) {
            double maxWidth = style.MaxWidth.Value;
            if (maxWidth <= 0 || double.IsNaN(maxWidth) || double.IsInfinity(maxWidth)) {
                throw new ArgumentException("Table max width must be a positive finite value.");
            }

            return Math.Min(frameWidth, maxWidth);
        }

        return frameWidth;
    }

    private static TableColumnLayout ResolveTableColumnLayout(TableBlock table, PdfOptions options, PdfTableStyle style, int columns, double frameWidth, double fontSize, int headerRowCount, int footerStartRowIndex, double padLeft, double padRight) {
        double[]? autoFitWeights = style.AutoFitColumns
            ? MeasureAutoFitColumnWeights(table, options, style, fontSize, headerRowCount, footerStartRowIndex)
            : null;
        double[]? autoFitMinimumWidths = style.AutoFitColumns
            ? MeasureAutoFitColumnMinimumWidths(table, options, style, fontSize, headerRowCount, footerStartRowIndex)
            : null;
        double tableWidth = ResolveTableAvailableWidth(style, frameWidth);
        double[] columnWidths = new double[columns];
        double[] columnWeights = new double[columns];
        bool[] fixedColumns = new bool[columns];
        double?[] minWidths = new double?[columns];
        double?[] maxWidths = new double?[columns];
        double fixedWidthTotal = 0D;
        double totalWeight = 0D;

        for (int column = 0; column < columns; column++) {
            double? minWidth = GetOptionalColumnWidth(style.ColumnMinWidthPoints, column, "Table minimum column widths must be positive finite values.");
            if (!minWidth.HasValue && autoFitMinimumWidths != null && column < autoFitMinimumWidths.Length) {
                minWidth = autoFitMinimumWidths[column];
            }

            double? maxWidth = GetOptionalColumnWidth(style.ColumnMaxWidthPoints, column, "Table maximum column widths must be positive finite values.");
            if (minWidth.HasValue && maxWidth.HasValue && minWidth.Value > maxWidth.Value + 0.001) {
                throw new ArgumentException("Table minimum column widths cannot be greater than maximum column widths.");
            }

            minWidths[column] = minWidth;
            maxWidths[column] = maxWidth;

            if (style.ColumnWidthPoints != null &&
                column < style.ColumnWidthPoints.Count &&
                style.ColumnWidthPoints[column].HasValue) {
                double fixedWidth = style.ColumnWidthPoints[column]!.Value;
                if (fixedWidth <= 0 || double.IsNaN(fixedWidth) || double.IsInfinity(fixedWidth)) {
                    throw new ArgumentException("Table fixed column widths must be positive finite values.");
                }

                if (minWidth.HasValue && fixedWidth < minWidth.Value - 0.001) {
                    throw new ArgumentException("Table fixed column widths cannot be smaller than configured minimum widths.");
                }

                if (maxWidth.HasValue && fixedWidth > maxWidth.Value + 0.001) {
                    throw new ArgumentException("Table fixed column widths cannot be greater than configured maximum widths.");
                }

                columnWidths[column] = fixedWidth;
                fixedColumns[column] = true;
                fixedWidthTotal += fixedWidth;
                continue;
            }

            double weight = 1D;
            if (style.ColumnWidthWeights != null && column < style.ColumnWidthWeights.Count) {
                weight = style.ColumnWidthWeights[column];
                if (weight <= 0 || double.IsNaN(weight) || double.IsInfinity(weight)) {
                    throw new ArgumentException("Table column width weights must be positive finite values.");
                }
            } else if (autoFitWeights != null && column < autoFitWeights.Length) {
                weight = autoFitWeights[column];
            }

            columnWeights[column] = weight;
            totalWeight += weight;
        }

        if (fixedWidthTotal > tableWidth + 0.001) {
            throw new ArgumentException("Table fixed column widths exceed the available table width.");
        }

        double remainingWidth = Math.Max(0D, tableWidth - fixedWidthTotal);
        if (totalWeight <= 0D) {
            remainingWidth = 0D;
        }

        DistributeFlexibleColumns(columnWidths, columnWeights, fixedColumns, minWidths, maxWidths, remainingWidth);
        tableWidth = Math.Min(tableWidth, columnWidths.Sum());
        for (int column = 0; column < columns; column++) {
            if (columnWidths[column] - padLeft - padRight <= 0.001) {
                throw new ArgumentException("Table horizontal cell padding must leave a positive text width.");
            }
        }

        return new TableColumnLayout {
            Widths = columnWidths,
            Width = tableWidth
        };
    }

    private static double ResolveTableX(PdfAlign align, PdfTableStyle style, double containerLeft, double containerWidth, double tableWidth) {
        double frameLeft = containerLeft + style.LeftIndent;
        double frameWidth = ResolveTableFrameWidth(style, containerWidth);
        if (align == PdfAlign.Center) {
            return frameLeft + Math.Max(0, (frameWidth - tableWidth) / 2);
        }

        if (align == PdfAlign.Right) {
            return frameLeft + Math.Max(0, frameWidth - tableWidth);
        }

        return frameLeft;
    }

    private static bool IsValidPdfAlign(PdfAlign align) =>
        align == PdfAlign.Left || align == PdfAlign.Center || align == PdfAlign.Right;

    private static bool IsValidPdfColumnAlign(PdfColumnAlign align) =>
        align == PdfColumnAlign.Left || align == PdfColumnAlign.Center || align == PdfColumnAlign.Right;

    private static bool IsValidPdfCellVerticalAlign(PdfCellVerticalAlign align) =>
        align == PdfCellVerticalAlign.Top || align == PdfCellVerticalAlign.Middle || align == PdfCellVerticalAlign.Bottom;

    private static OfficeIMO.Drawing.OfficeFontInfo ToOfficeFontInfo(PdfStandardFont font, double size) {
        string family = font switch {
            PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBoldItalic => "Times New Roman",
            PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique => "Courier New",
            _ => "Helvetica"
        };

        OfficeIMO.Drawing.OfficeFontStyle style = OfficeIMO.Drawing.OfficeFontStyle.Regular;
        switch (font) {
            case PdfStandardFont.HelveticaBold:
            case PdfStandardFont.HelveticaBoldOblique:
            case PdfStandardFont.TimesBold:
            case PdfStandardFont.TimesBoldItalic:
            case PdfStandardFont.CourierBold:
            case PdfStandardFont.CourierBoldOblique:
                style |= OfficeIMO.Drawing.OfficeFontStyle.Bold;
                break;
        }

        switch (font) {
            case PdfStandardFont.HelveticaOblique:
            case PdfStandardFont.HelveticaBoldOblique:
            case PdfStandardFont.TimesItalic:
            case PdfStandardFont.TimesBoldItalic:
            case PdfStandardFont.CourierOblique:
            case PdfStandardFont.CourierBoldOblique:
                style |= OfficeIMO.Drawing.OfficeFontStyle.Italic;
                break;
        }

        return new OfficeIMO.Drawing.OfficeFontInfo(family, size, style);
    }

    private static double[] MeasureAutoFitColumnWeights(TableBlock table, PdfOptions options, PdfTableStyle style, double fontSize, int headerRowCount, int footerStartRowIndex) {
        int cols = GetTableColumnCount(table);
        var weights = new double[cols];
        var normalFont = ToOfficeFontInfo(ChooseNormal(options.DefaultFont), fontSize);
        var measurer = OfficeIMO.Drawing.OfficeTextMeasurer.Create(normalFont);

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            double rowSize = GetTableRowFontSize(style, rowIndex, headerRowCount, footerStartRowIndex, fontSize);
            var rowFont = ToOfficeFontInfo(GetTableRowFont(options, GetTableRowBold(style, rowIndex, headerRowCount, footerStartRowIndex)), rowSize);
            var measurementStyle = measurer.CreateStyle(rowFont);
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                double measuredPoints = measurer.MeasureWidth(cell.Text, measurementStyle) * 72D / measurementStyle.Dpi;
                double requestedWidth = Math.Max(1D, measuredPoints + GetTableCellPaddingLeft(style) + GetTableCellPaddingRight(style));
                double requestedPerColumn = requestedWidth / cell.ColumnSpan;
                for (int c = cell.Column; c < cell.Column + cell.ColumnSpan && c < cols; c++) {
                    if (requestedPerColumn > weights[c]) {
                        weights[c] = requestedPerColumn;
                    }
                }
            }
        }

        for (int c = 0; c < weights.Length; c++) {
            if (weights[c] <= 0D) {
                weights[c] = 1D;
            }
        }

        return weights;
    }

    private static double[] MeasureAutoFitColumnMinimumWidths(TableBlock table, PdfOptions options, PdfTableStyle style, double fontSize, int headerRowCount, int footerStartRowIndex) {
        int cols = GetTableColumnCount(table);
        var widths = new double[cols];
        double maximumTokenWidth = Math.Max(1D, fontSize * 12D);

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            double rowSize = GetTableRowFontSize(style, rowIndex, headerRowCount, footerStartRowIndex, fontSize);
            var rowFont = GetTableRowFont(options, GetTableRowBold(style, rowIndex, headerRowCount, footerStartRowIndex));
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                double tokenWidth = 0D;
                string[] tokens = cell.Text
                    .Replace("\r\n", "\n")
                    .Replace('\r', '\n')
                    .Split(TokenSplitChars, StringSplitOptions.RemoveEmptyEntries);
                if (tokens.Length == 0) {
                    tokenWidth = EstimateSimpleTextWidth(cell.Text, rowFont, rowSize);
                } else {
                    for (int tokenIndex = 0; tokenIndex < tokens.Length; tokenIndex++) {
                        tokenWidth = Math.Max(tokenWidth, EstimateSimpleTextWidth(tokens[tokenIndex], rowFont, rowSize));
                    }
                }

                double requestedWidth = Math.Max(1D, Math.Min(tokenWidth, maximumTokenWidth) + GetTableCellPaddingLeft(style) + GetTableCellPaddingRight(style));
                double requestedPerColumn = requestedWidth / cell.ColumnSpan;
                for (int columnIndex = cell.Column; columnIndex < cell.Column + cell.ColumnSpan && columnIndex < cols; columnIndex++) {
                    if (requestedPerColumn > widths[columnIndex]) {
                        widths[columnIndex] = requestedPerColumn;
                    }
                }
            }
        }

        for (int columnIndex = 0; columnIndex < widths.Length; columnIndex++) {
            if (widths[columnIndex] <= 0D) {
                widths[columnIndex] = 1D;
            }
        }

        return widths;
    }

    private static PageImage CreatePageImage(ImageBlock block, PdfImageStyle style, double targetX, double targetBottomY) {
        double drawX = targetX;
        double drawY = targetBottomY;
        double drawWidth = block.Width;
        double drawHeight = block.Height;
        OfficeClipPath? clipPath = style.ClipPath?.Clone();

        if (style.Fit != OfficeImageFit.Stretch) {
            double imageAspect = block.Info.Width / (double)block.Info.Height;
            double targetAspect = block.Width / block.Height;

            if (style.Fit == OfficeImageFit.Contain) {
                if (targetAspect > imageAspect) {
                    drawHeight = block.Height;
                    drawWidth = drawHeight * imageAspect;
                    drawX = targetX + (block.Width - drawWidth) / 2D;
                } else {
                    drawWidth = block.Width;
                    drawHeight = drawWidth / imageAspect;
                    drawY = targetBottomY + (block.Height - drawHeight) / 2D;
                }
            } else {
                if (targetAspect > imageAspect) {
                    drawWidth = block.Width;
                    drawHeight = drawWidth / imageAspect;
                    drawY = targetBottomY + (block.Height - drawHeight) / 2D;
                } else {
                    drawHeight = block.Height;
                    drawWidth = drawHeight * imageAspect;
                    drawX = targetX + (block.Width - drawWidth) / 2D;
                }

                if (clipPath == null) {
                    clipPath = OfficeClipPath.Rectangle(block.Width, block.Height);
                }
            }
        }

        return new PageImage {
            Data = block.Data,
            Info = block.Info,
            X = drawX,
            Y = drawY,
            W = drawWidth,
            H = drawHeight,
            ClipPath = clipPath,
            ClipX = targetX,
            ClipY = targetBottomY,
            ClipHeight = block.Height
        };
    }

    private static void DistributeFlexibleColumns(
        double[] widths,
        double[] weights,
        bool[] fixedColumns,
        double?[] minWidths,
        double?[] maxWidths,
        double availableWidth) {
        var active = new bool[widths.Length];
        int activeCount = 0;
        double requiredMinimum = 0;

        for (int i = 0; i < widths.Length; i++) {
            if (fixedColumns[i]) {
                continue;
            }

            active[i] = true;
            activeCount++;
            if (minWidths[i].HasValue) {
                requiredMinimum += minWidths[i]!.Value;
            }
        }

        if (requiredMinimum > availableWidth + 0.001) {
            throw new ArgumentException("Table minimum column widths exceed the available table width.");
        }

        double remaining = availableWidth;
        while (activeCount > 0) {
            double weightSum = 0;
            for (int i = 0; i < weights.Length; i++) {
                if (active[i]) {
                    weightSum += weights[i];
                }
            }

            bool constrained = false;
            for (int i = 0; i < widths.Length; i++) {
                if (!active[i]) {
                    continue;
                }

                double proposed = weightSum > 0 ? remaining * (weights[i] / weightSum) : remaining / activeCount;
                if (minWidths[i].HasValue && proposed < minWidths[i]!.Value) {
                    widths[i] = minWidths[i]!.Value;
                    remaining -= widths[i];
                    active[i] = false;
                    activeCount--;
                    constrained = true;
                } else if (maxWidths[i].HasValue && proposed > maxWidths[i]!.Value) {
                    widths[i] = maxWidths[i]!.Value;
                    remaining -= widths[i];
                    active[i] = false;
                    activeCount--;
                    constrained = true;
                }
            }

            if (constrained) {
                continue;
            }

            for (int i = 0; i < widths.Length; i++) {
                if (!active[i]) {
                    continue;
                }

                widths[i] = weightSum > 0 ? remaining * (weights[i] / weightSum) : remaining / activeCount;
                active[i] = false;
            }
            break;
        }
    }


    private static LayoutResult LayoutBlocks(IEnumerable<IPdfBlock> blocks, PdfOptions opts) {
        var sb = new StringBuilder();
        var pages = new System.Collections.Generic.List<LayoutResult.Page>();
        var optionsStack = new System.Collections.Generic.Stack<PdfOptions>();
        optionsStack.Push(opts);
        var pageGroupStack = new System.Collections.Generic.Stack<int>();
        pageGroupStack.Push(0);
        PdfOptions currentOpts = opts;
        int currentPageGroupId = 0;
        int nextPageGroupId = 1;

        LayoutResult.Page? currentPage = null;
        double width = 0;
        double yStart = 0;
        double y = 0;
        bool pageDirty = false;

        bool usedBold = false;
        bool usedItalic = false;
        bool usedBoldItalic = false;

        void StartPage(PdfOptions options) {
            options.Validate();
            currentOpts = options;
            width = options.PageWidth - options.MarginLeft - options.MarginRight;
            yStart = options.PageHeight - options.MarginTop;
            y = yStart;
            currentPage = new LayoutResult.Page { Options = options, PageGroupId = currentPageGroupId };
            sb.Clear();
            pageDirty = false;
        }

        void EnsurePage() {
            if (currentPage == null) StartPage(currentOpts);
        }

        void FlushPage(bool force = false) {
            if (currentPage == null) return;
            if (!force && !pageDirty && currentPage.Images.Count == 0 && currentPage.Annotations.Count == 0 && currentPage.GraphicsStates.Count == 0 && currentPage.Shadings.Count == 0 && currentPage.NamedDestinations.Count == 0) {
                currentPage = null;
                sb.Clear();
                pageDirty = false;
                return;
            }
            currentPage.Content = sb.ToString();
            pages.Add(currentPage);
            currentPage = null;
            sb.Clear();
            pageDirty = false;
        }

        void NewPage() {
            FlushPage(pageDirty || (currentPage?.Images.Count ?? 0) > 0 || (currentPage?.Annotations.Count ?? 0) > 0 || (currentPage?.GraphicsStates.Count ?? 0) > 0 || (currentPage?.Shadings.Count ?? 0) > 0 || (currentPage?.NamedDestinations.Count ?? 0) > 0);
            StartPage(currentOpts);
        }

        void WriteLinesInternal(string fontRes, double fontSize, double lineHeight, double x, double widthUsed, double startY, System.Collections.Generic.IReadOnlyList<string> lines, PdfAlign align, PdfColor? color = null, bool applyBaselineTweak = false) {
            EnsurePage();
            pageDirty = true;
            var content = new ContentStreamBuilder(sb)
                .BeginText()
                .Font(fontRes, fontSize)
                .TextLeading(lineHeight);
            var lineFont = fontRes == "F2" ? ChooseBold(ChooseNormal(currentOpts.DefaultFont)) : ChooseNormal(currentOpts.DefaultFont);
            double yStart2 = startY;
            if (applyBaselineTweak) {
                yStart2 -= GetDescender(lineFont, fontSize) * 0.0;
            }
            content.TextMatrix(x, yStart2);
            var effectiveColor = color ?? currentOpts.DefaultTextColor;
            if (effectiveColor.HasValue) content.FillColor(effectiveColor.Value);
            for (int i = 0; i < lines.Count; i++) {
                string line = lines[i];
                double dx = 0;
                double estWidth = EstimateSimpleTextWidth(line, lineFont, fontSize);
                if (align == PdfAlign.Center) dx = Math.Max(0, (widthUsed - estWidth) / 2);
                else if (align == PdfAlign.Right) dx = Math.Max(0, (widthUsed - estWidth));
                if (Math.Abs(dx) > 0.0001) content.MoveText(dx, 0);
                content.ShowHexText(EncodeWinAnsiHex(line));
                if (Math.Abs(dx) > 0.0001) content.MoveText(-dx, 0);
                if (i != lines.Count - 1) content.NextTextLine();
            }
            content.EndText();
        }

        void WriteLines(string fontRes, double fontSize, double lineHeight, double x, double startY, System.Collections.Generic.IReadOnlyList<string> lines, PdfAlign align, PdfColor? color = null, bool applyBaselineTweak = false)
            => WriteLinesInternal(fontRes, fontSize, lineHeight, x, width, startY, lines, align, color, applyBaselineTweak);

        void AddHeadingLinkAnnotations(HeadingBlock heading, System.Collections.Generic.IReadOnlyList<string> lines, PdfStandardFont font, double fontSize, double lineHeight, double x, double widthUsed, double startBaselineY) {
            if (string.IsNullOrEmpty(heading.LinkUri)) {
                return;
            }

            double asc = GetAscender(font, fontSize);
            double desc = GetDescender(font, fontSize);
            for (int i = 0; i < lines.Count; i++) {
                string line = lines[i];
                double lineWidth = EstimateSimpleTextWidth(line, font, fontSize);
                if (lineWidth <= 0.001D) {
                    continue;
                }

                double dx = 0D;
                if (heading.Align == PdfAlign.Center) dx = Math.Max(0, (widthUsed - lineWidth) / 2);
                else if (heading.Align == PdfAlign.Right) dx = Math.Max(0, widthUsed - lineWidth);
                double baselineY = startBaselineY - i * lineHeight;
                double x1 = x + dx;
                double x2 = x1 + Math.Min(widthUsed, lineWidth);
                double y1 = baselineY - desc;
                double y2 = baselineY + asc;
                currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = heading.LinkUri!, Contents = heading.LinkContents });
            }
        }

        void AddImageLinkAnnotation(ImageBlock image, PdfImageStyle style, PageImage pageImage, double targetX, double targetBottomY) {
            if (string.IsNullOrEmpty(image.LinkUri)) {
                return;
            }

            double x1 = pageImage.X;
            double y1 = pageImage.Y;
            double x2 = pageImage.X + pageImage.W;
            double y2 = pageImage.Y + pageImage.H;
            if (style.Fit == OfficeImageFit.Cover || style.ClipPath != null) {
                x1 = targetX;
                y1 = targetBottomY;
                x2 = targetX + image.Width;
                y2 = targetBottomY + image.Height;
            }

            currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = image.LinkUri!, Contents = image.LinkContents });
        }

        double GetAlignedObjectX(double containerX, double containerWidth, double objectWidth, PdfAlign align) {
            if (align == PdfAlign.Center) return containerX + Math.Max(0, (containerWidth - objectWidth) / 2);
            if (align == PdfAlign.Right) return containerX + Math.Max(0, containerWidth - objectWidth);
            return containerX;
        }

        void AddShapeLinkAnnotation(ShapeBlock shape, PdfDrawingStyle style, double containerX, double containerWidth, double topY) {
            if (string.IsNullOrEmpty(shape.LinkUri)) {
                return;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, shape.Shape.Width, style.Align);
            currentPage!.Annotations.Add(new LinkAnnotation { X1 = x, Y1 = topY - shape.Shape.Height, X2 = x + shape.Shape.Width, Y2 = topY, Uri = shape.LinkUri!, Contents = shape.LinkContents });
        }

        void AddDrawingLinkAnnotation(DrawingBlock drawing, PdfDrawingStyle style, double containerX, double containerWidth, double topY) {
            if (string.IsNullOrEmpty(drawing.LinkUri)) {
                return;
            }

            double x = GetAlignedObjectX(containerX, containerWidth, drawing.Drawing.Width, style.Align);
            currentPage!.Annotations.Add(new LinkAnnotation { X1 = x, Y1 = topY - drawing.Drawing.Height, X2 = x + drawing.Drawing.Width, Y2 = topY, Uri = drawing.LinkUri!, Contents = drawing.LinkContents });
        }

        void AddNamedDestination(BookmarkBlock bookmark, double topY) {
            EnsurePage();
            currentPage!.NamedDestinations.Add(new PageNamedDestination { Name = bookmark.Name, Y = topY });
        }

        double FirstTextBaselineFromTop(PdfStandardFont font, double fontSize, double topY) =>
            topY - GetAscender(font, fontSize);

        PdfColor? ToPdfColor(OfficeIMO.Drawing.OfficeColor? color) =>
            color.HasValue ? PdfColor.FromOfficeColorOrNull(color.Value) : null;

        string? EnsureGraphicsState(double fillOpacity, double strokeOpacity) {
            if (fillOpacity >= 1D && strokeOpacity >= 1D) {
                return null;
            }

            EnsurePage();
            for (int i = 0; i < currentPage!.GraphicsStates.Count; i++) {
                var existing = currentPage.GraphicsStates[i];
                if (existing.FillOpacity.Equals(fillOpacity) && existing.StrokeOpacity.Equals(strokeOpacity)) {
                    return existing.Name;
                }
            }

            string name = "GS" + (currentPage.GraphicsStates.Count + 1).ToString(CultureInfo.InvariantCulture);
            currentPage.GraphicsStates.Add(new PageGraphicsState {
                Name = name,
                FillOpacity = fillOpacity,
                StrokeOpacity = strokeOpacity
            });
            return name;
        }

        string? EnsureOpacityState(OfficeIMO.Drawing.OfficeShape shape) {
            bool hasFill = (shape.FillColor.HasValue || shape.FillGradient != null) && shape.Kind != OfficeIMO.Drawing.OfficeShapeKind.Line;
            bool hasStroke = shape.StrokeColor.HasValue && shape.StrokeWidth > 0;
            double fillOpacity = hasFill ? shape.FillOpacity ?? 1D : 1D;
            double strokeOpacity = hasStroke ? shape.StrokeOpacity ?? 1D : 1D;
            return EnsureGraphicsState(fillOpacity, strokeOpacity);
        }

        string? EnsureLinearGradient(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY, bool localCoordinates) {
            var gradient = shape.FillGradient;
            if (gradient == null || shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) {
                return null;
            }

            var start = gradient.Stops[0].Color;
            var end = gradient.Stops[1].Color;
            double originX = localCoordinates ? 0D : xShape;
            double originY = localCoordinates ? 0D : bottomY;
            double x0 = originX + gradient.StartX * shape.Width;
            double y0 = originY + shape.Height - gradient.StartY * shape.Height;
            double x1 = originX + gradient.EndX * shape.Width;
            double y1 = originY + shape.Height - gradient.EndY * shape.Height;

            EnsurePage();
            for (int i = 0; i < currentPage!.Shadings.Count; i++) {
                var existing = currentPage.Shadings[i];
                if (existing.StartColor.Equals(start) &&
                    existing.EndColor.Equals(end) &&
                    existing.X0.Equals(x0) &&
                    existing.Y0.Equals(y0) &&
                    existing.X1.Equals(x1) &&
                    existing.Y1.Equals(y1)) {
                    return existing.Name;
                }
            }

            string name = "SH" + (currentPage.Shadings.Count + 1).ToString(CultureInfo.InvariantCulture);
            currentPage.Shadings.Add(new PageShading {
                Name = name,
                StartColor = start,
                EndColor = end,
                X0 = x0,
                Y0 = y0,
                X1 = x1,
                Y1 = y1
            });
            return name;
        }

        void DrawShapeShadowAt(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY) {
            var shadow = shape.Shadow;
            if (shadow == null || shadow.Opacity <= 0D) {
                return;
            }

            PdfColor shadowColor = PdfColor.FromOfficeColor(shadow.Color);
            double shadowX = xShape + shadow.OffsetX;
            double shadowBottomY = bottomY - shadow.OffsetY;
            string? shadowState = EnsureGraphicsState(shadow.Opacity, shadow.Opacity);

            var content = new ContentStreamBuilder(sb)
                .SaveState();
            if (shadowState != null) {
                content.GraphicsState(shadowState);
            }

            if (shape.Transform.HasValue) {
                DrawTransformedShape(
                    sb,
                    shape,
                    shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line ? null : shadowColor,
                    shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line ? shadowColor : null,
                    null,
                    shadowX,
                    shadowBottomY);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) {
                DrawLine(sb, shadowColor, shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, shadowX, shadowBottomY, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.RoundedRectangle) {
                DrawRoundedRectangle(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height, shape.CornerRadius);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Rectangle) {
                DrawRectangle(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Ellipse) {
                DrawEllipse(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shadowX, shadowBottomY, shape.Width, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Polygon) {
                DrawPolygon(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, shadowX, shadowBottomY, shape.Height);
            } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Path) {
                DrawPath(sb, shadowColor, null, 0, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.PathCommands, shadowX, shadowBottomY, shape.Height);
            }

            content.RestoreState();
            pageDirty = true;
        }

        void DrawShapeGeometryAt(OfficeIMO.Drawing.OfficeShape shape, double xShape, double bottomY) {
            DrawShapeShadowAt(shape, xShape, bottomY);

            string? opacityState = EnsureOpacityState(shape);
            if (opacityState != null) {
                new ContentStreamBuilder(sb)
                    .SaveState()
                    .GraphicsState(opacityState);
            }

            if (shape.Transform.HasValue) {
                pageDirty = true;
                string? shadingName = EnsureLinearGradient(shape, xShape, bottomY, localCoordinates: true);
                DrawTransformedShape(sb, shape, shadingName == null ? ToPdfColor(shape.FillColor) : null, ToPdfColor(shape.StrokeColor), shadingName, xShape, bottomY);
            } else {
                if (shape.ClipPath != null) {
                    new ContentStreamBuilder(sb)
                        .SaveState();
                    AppendClipPath(sb, shape.ClipPath, xShape, bottomY, shape.Height);
                }

                string? shadingName = EnsureLinearGradient(shape, xShape, bottomY, localCoordinates: false);
                if (shadingName != null) {
                    pageDirty = true;
                    DrawGradientShape(sb, shape, shadingName, xShape, bottomY);
                }

                PdfColor? fillColor = shadingName == null ? ToPdfColor(shape.FillColor) : null;
                if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Line) {
                    pageDirty = true;
                    DrawLine(sb, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, xShape, bottomY, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.RoundedRectangle) {
                    pageDirty = true;
                    DrawRoundedRectangle(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height, shape.CornerRadius);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Rectangle) {
                    pageDirty = true;
                    DrawRectangle(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Ellipse) {
                    pageDirty = true;
                    DrawEllipse(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, xShape, bottomY, shape.Width, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Polygon) {
                    pageDirty = true;
                    DrawPolygon(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.Points, xShape, bottomY, shape.Height);
                } else if (shape.Kind == OfficeIMO.Drawing.OfficeShapeKind.Path) {
                    pageDirty = true;
                    DrawPath(sb, fillColor, ToPdfColor(shape.StrokeColor), shape.StrokeWidth, shape.StrokeDashStyle, shape.StrokeLineCap, shape.StrokeLineJoin, shape.PathCommands, xShape, bottomY, shape.Height);
                }

                if (shape.ClipPath != null) {
                    new ContentStreamBuilder(sb)
                        .RestoreState();
                }
            }

            if (opacityState != null) {
                new ContentStreamBuilder(sb)
                    .RestoreState();
            }
        }

        void DrawShapeAt(ShapeBlock block, PdfDrawingStyle style, double containerX, double containerWidth, double topY) {
            double xShape = GetAlignedObjectX(containerX, containerWidth, block.Shape.Width, style.Align);
            DrawShapeGeometryAt(block.Shape, xShape, topY - block.Shape.Height);
        }

        void DrawDrawingAt(DrawingBlock block, PdfDrawingStyle style, double containerX, double containerWidth, double topY) {
            double xDrawing = GetAlignedObjectX(containerX, containerWidth, block.Drawing.Width, style.Align);
            for (int i = 0; i < block.Drawing.Shapes.Count; i++) {
                var item = block.Drawing.Shapes[i];
                double xShape = xDrawing + item.X;
                double bottomY = topY - item.Y - item.Shape.Height;
                DrawShapeGeometryAt(item.Shape, xShape, bottomY);
            }
        }

        void RenderShapeBlock(ShapeBlock block, double containerX, double containerWidth) {
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDoc.ValidateDrawingStyle(style, "Shape");
            double needed = style.SpacingBefore + block.Shape.Height + style.SpacingAfter;
            EnsureFixedFlowBlockFits("Shape", block.Shape.Width, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) { NewPage(); }
            y -= style.SpacingBefore;
            DrawShapeAt(block, style, containerX, containerWidth, y);
            AddShapeLinkAnnotation(block, style, containerX, containerWidth, y);
            y -= block.Shape.Height + style.SpacingAfter;
        }

        void RenderDrawingBlock(DrawingBlock block, double containerX, double containerWidth) {
            PdfDrawingStyle style = ResolveDrawingStyle(block, currentOpts);
            PdfDoc.ValidateDrawingStyle(style, "Drawing");
            double needed = style.SpacingBefore + block.Drawing.Height + style.SpacingAfter;
            EnsureFixedFlowBlockFits("Drawing", block.Drawing.Width, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) { NewPage(); }
            y -= style.SpacingBefore;
            DrawDrawingAt(block, style, containerX, containerWidth, y);
            AddDrawingLinkAnnotation(block, style, containerX, containerWidth, y);
            y -= block.Drawing.Height + style.SpacingAfter;
        }

        void KeepFixedBlockWithNext(double needed, double nextHeight) {
            double keepHeight = needed + nextHeight;
            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                NewPage();
            }
        }

        void RenderHorizontalRuleBlock(HorizontalRuleBlock block, double containerX, double containerWidth) {
            PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(block, currentOpts);
            ValidateHorizontalRule(ruleStyle);
            double needed = ruleStyle.SpacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
            EnsureFixedFlowBlockFits("Horizontal rule", containerWidth, needed, containerWidth);
            if (y - needed < currentOpts.MarginBottom) { NewPage(); }
            y -= ruleStyle.SpacingBefore;
            double yLine = y - ruleStyle.Thickness * 0.5;
            DrawHLine(sb, ruleStyle.Color, ruleStyle.Thickness, containerX, containerX + containerWidth, yLine);
            pageDirty = true;
            y -= ruleStyle.Thickness + ruleStyle.SpacingAfter;
        }

        void EnsureFixedFlowBlockFits(string blockName, double blockWidth, double blockHeight, double availableWidth) {
            if (blockWidth > availableWidth + 0.001) {
                throw new ArgumentException(blockName + " width exceeds the available page content width.");
            }

            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            if (blockHeight > availableHeight + 0.001) {
                throw new ArgumentException(blockName + " height exceeds the available page content height.");
            }
        }

        void ValidateHorizontalRule(PdfHorizontalRuleStyle rule) {
            if (rule.Thickness <= 0 || double.IsNaN(rule.Thickness) || double.IsInfinity(rule.Thickness)) {
                throw new ArgumentException("Horizontal rule thickness must be a positive finite value.");
            }

            if (rule.SpacingBefore < 0 || double.IsNaN(rule.SpacingBefore) || double.IsInfinity(rule.SpacingBefore)) {
                throw new ArgumentException("Horizontal rule spacing before must be a non-negative finite value.");
            }

            if (rule.SpacingAfter < 0 || double.IsNaN(rule.SpacingAfter) || double.IsInfinity(rule.SpacingAfter)) {
                throw new ArgumentException("Horizontal rule spacing after must be a non-negative finite value.");
            }
        }

        void ValidatePanelStyle(PanelStyle style, double panelWidth) {
            Guard.LeftCenterRightAlign(style.Align, nameof(style.Align), "Panel box");

            if (style.BorderWidth < 0 || double.IsNaN(style.BorderWidth) || double.IsInfinity(style.BorderWidth)) {
                throw new ArgumentException("Panel border width must be a non-negative finite value.");
            }

            if (style.PaddingX < 0 || double.IsNaN(style.PaddingX) || double.IsInfinity(style.PaddingX)) {
                throw new ArgumentException("Panel horizontal padding must be a non-negative finite value.");
            }

            if (style.PaddingY < 0 || double.IsNaN(style.PaddingY) || double.IsInfinity(style.PaddingY)) {
                throw new ArgumentException("Panel vertical padding must be a non-negative finite value.");
            }

            if (style.MaxWidth.HasValue && (style.MaxWidth.Value <= 0 || double.IsNaN(style.MaxWidth.Value) || double.IsInfinity(style.MaxWidth.Value))) {
                throw new ArgumentException("Panel maximum width must be a positive finite value.");
            }

            if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
                throw new ArgumentException("Panel spacing before must be a non-negative finite value.");
            }

            if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
                throw new ArgumentException("Panel spacing after must be a non-negative finite value.");
            }

            if (panelWidth - 2 * style.PaddingX <= 0) {
                throw new ArgumentException("Panel horizontal padding must leave a positive text width.");
            }
        }

        void RenderListItem(System.Collections.Generic.List<string> lines, string marker, double markerX, double markerWidth, PdfAlign markerAlign, double textX, double textWidth, PdfAlign textAlign, PdfColor? color, double size, double leading, double spacingBefore, double spacingAfter) {
            int lineIndex = 0;
            bool firstSegment = true;
            var listFont = ChooseNormal(currentOpts.DefaultFont);
            if (spacingBefore > 0) {
                if (y - spacingBefore < currentOpts.MarginBottom) {
                    NewPage();
                }

                y -= spacingBefore;
            }

            while (lineIndex < lines.Count) {
                double available = y - currentOpts.MarginBottom;
                if (available < leading) {
                    NewPage();
                    available = y - currentOpts.MarginBottom;
                    if (available < leading) {
                        break;
                    }
                }

                int take = 0;
                double heightSum = 0;
                for (int k = lineIndex; k < lines.Count; k++) {
                    if (heightSum + leading > available) {
                        break;
                    }

                    heightSum += leading;
                    take++;
                }

                if (take == 0) {
                    NewPage();
                    continue;
                }

                var segmentLines = new System.Collections.Generic.List<string>(take);
                for (int k = 0; k < take; k++) {
                    segmentLines.Add(lines[lineIndex + k]);
                }

                double baselineY = FirstTextBaselineFromTop(listFont, size, y);
                if (firstSegment) {
                    var markerLines = new System.Collections.Generic.List<string>(1) { marker };
                    WriteLinesInternal("F1", size, leading, markerX, markerWidth, baselineY, markerLines, markerAlign, color, applyBaselineTweak: true);
                }

                WriteLinesInternal("F1", size, leading, textX, textWidth, baselineY, segmentLines, textAlign, color, applyBaselineTweak: true);
                y -= heightSum;
                lineIndex += take;
                firstSegment = false;
                if (lineIndex < lines.Count) {
                    NewPage();
                } else {
                    y -= spacingAfter;
                }
            }
        }

        double MeasureListKeepTogetherHeight(System.Collections.Generic.IReadOnlyList<System.Collections.Generic.List<string>> itemLines, double leading, double spacingBefore, double itemSpacing, double spacingAfter) {
            double total = 0D;
            for (int itemIndex = 0; itemIndex < itemLines.Count; itemIndex++) {
                total += itemIndex == 0 ? spacingBefore : 0D;
                total += itemLines[itemIndex].Count * leading;
                total += itemIndex == itemLines.Count - 1 ? spacingAfter : itemSpacing;
            }

            return total;
        }

        PdfParagraphStyle? EffectiveParagraphStyle(RichParagraphBlock paragraph) => paragraph.Style ?? currentOpts.DefaultParagraphStyleSnapshot;

        double MeasureNextParagraphFirstLineHeight(RichParagraphBlock paragraph, double frameX, double frameWidth, double fontSize) {
            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph);
            double leading = GetParagraphLeading(paragraphStyle, fontSize);
            double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
            var textFrame = GetParagraphTextFrame(paragraphStyle, frameX, frameWidth);
            var wrap = WrapRichRuns(paragraph.Runs, textFrame.Width, fontSize, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, GetParagraphTabStopWidth(paragraphStyle));
            return wrap.LineHeights.Count == 0 ? spacingBefore : spacingBefore + wrap.LineHeights[0];
        }

        double MeasureNextBlockFirstVisualHeight(IPdfBlock block, double frameX, double frameWidth, double fontSize) {
            if (block is RichParagraphBlock paragraph) {
                return MeasureNextParagraphFirstLineHeight(paragraph, frameX, frameWidth, fontSize);
            }

            if (block is HeadingBlock heading) {
                PdfHeadingStyle? headingStyle = ResolveHeadingStyle(heading, currentOpts);
                double headingSize = GetHeadingFontSize(heading, headingStyle);
                double headingLeading = GetHeadingLeading(headingStyle, headingSize);
                return (headingStyle?.SpacingBefore ?? 0D) + headingLeading;
            }

            if (block is SpacerBlock spacer) {
                return spacer.Height;
            }

            if (block is BulletListBlock bullets) {
                PdfListStyle? listStyle = ResolveListStyle(bullets, currentOpts);
                double size = GetListFontSize(listStyle, fontSize);
                double leading = GetListLeading(listStyle, size);
                string? firstItem = bullets.Items.Count > 0 ? bullets.Items[0] : null;
                if (firstItem == null) {
                    return listStyle?.SpacingBefore ?? 0D;
                }

                return (listStyle?.SpacingBefore ?? 0D) + leading;
            }

            if (block is NumberedListBlock numbered) {
                PdfListStyle? listStyle = ResolveListStyle(numbered, currentOpts);
                double size = GetListFontSize(listStyle, fontSize);
                double leading = GetListLeading(listStyle, size);
                string? firstItem = numbered.Items.Count > 0 ? numbered.Items[0] : null;
                if (firstItem == null) {
                    return listStyle?.SpacingBefore ?? 0D;
                }

                return (listStyle?.SpacingBefore ?? 0D) + leading;
            }

            if (block is PanelParagraphBlock panel) {
                PanelStyle panelStyle = ResolvePanelStyle(panel, currentOpts);
                double innerWidth = panelStyle.MaxWidth.HasValue ? Math.Min(frameWidth, panelStyle.MaxWidth.Value) : frameWidth;
                ValidatePanelStyle(panelStyle, innerWidth);
                double size = fontSize;
                double leading = size * 1.4;
                double textWidth = innerWidth - 2 * panelStyle.PaddingX;
                var wrap = WrapRichRuns(panel.Runs, textWidth, size, ChooseNormal(currentOpts.DefaultFont), leading);
                double firstLineHeight = wrap.LineHeights.Count == 0 ? 0D : wrap.LineHeights[0];
                return panelStyle.SpacingBefore + panelStyle.PaddingY + firstLineHeight + panelStyle.PaddingY;
            }

            if (block is TableBlock table) {
                PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
                int columns = GetTableColumnCount(table);
                if (columns == 0) {
                    return style.SpacingBefore;
                }

                double padLeft = GetTableCellPaddingLeft(style);
                double padRight = GetTableCellPaddingRight(style);
                double padTop = GetTableCellPaddingTop(style);
                double padBottom = GetTableCellPaddingBottom(style);
                ValidateTableRoleRowCounts(style, table.Rows.Count);
                int headerRowCount = style.HeaderRowCount;
                int footerRowCount = style.FooterRowCount;
                int footerStartRowIndex = table.Rows.Count - footerRowCount;
                ValidateTableCellStyleCoordinates(style, table.Rows.Count, columns);
                ValidateTableColumnStyleBounds(style, columns);
                ValidateTableRowSpansWithinRoleBoundaries(table, columns, headerRowCount, footerStartRowIndex);
                double tableFontSize = GetTableBodyFontSize(style, fontSize);
                TableColumnLayout columnLayout = ResolveTableColumnLayout(table, currentOpts, style, columns, frameWidth, tableFontSize, headerRowCount, footerStartRowIndex, padLeft, padRight);
                double tableWidth = columnLayout.Width;
                double rowSize = GetTableRowFontSize(style, 0, headerRowCount, footerStartRowIndex, fontSize);
                double rowLeading = GetTableLeading(style, rowSize);
                bool rowUsesBold = GetTableRowBold(style, 0, headerRowCount, footerStartRowIndex);
                int maxLines = 1;
                var firstRowCells = GetTableCellLayouts(table, 0, columns);
                for (int cellIndex = 0; cellIndex < firstRowCells.Count; cellIndex++) {
                    TableCellLayout cell = firstRowCells[cellIndex];
                    double cellWidth = GetTableCellWidth(columnLayout.Widths, cell.Column, cell.ColumnSpan, 0D);
                    double innerWidth = Math.Max(1D, cellWidth - padLeft - padRight);
                    var lines = WrapSimpleText(cell.Text, innerWidth, GetTableRowFont(currentOpts, rowUsesBold), rowSize);
                    maxLines = Math.Max(maxLines, lines.Count);
                }

                double firstRowHeight = Math.Max(maxLines * rowLeading + padTop + padBottom, style.MinRowHeight);
                double captionHeight = 0D;
                if (!string.IsNullOrWhiteSpace(style.Caption)) {
                    double captionSize = style.CaptionFontSize ?? fontSize;
                    double captionLeading = captionSize * 1.25D;
                    var captionLines = WrapSimpleText(style.Caption!, tableWidth, ChooseNormal(currentOpts.DefaultFont), captionSize);
                    captionHeight = captionLines.Count * captionLeading + style.CaptionSpacingAfter;
                }

                return style.SpacingBefore + captionHeight + firstRowHeight;
            }

            if (block is HorizontalRuleBlock rule) {
                PdfHorizontalRuleStyle style = ResolveHorizontalRuleStyle(rule, currentOpts);
                return style.SpacingBefore + style.Thickness + style.SpacingAfter;
            }

            if (block is ImageBlock image) {
                PdfImageStyle style = ResolveImageStyle(image, currentOpts);
                return style.SpacingBefore + image.Height + style.SpacingAfter;
            }

            if (block is ShapeBlock shape) {
                PdfDrawingStyle style = ResolveDrawingStyle(shape, currentOpts);
                return style.SpacingBefore + shape.Shape.Height + style.SpacingAfter;
            }

            if (block is DrawingBlock drawing) {
                PdfDrawingStyle style = ResolveDrawingStyle(drawing, currentOpts);
                return style.SpacingBefore + drawing.Drawing.Height + style.SpacingAfter;
            }

            if (block is RowBlock row) {
                int columns = row.Columns.Count;
                if (columns == 0) {
                    return 0D;
                }

                PdfRowStyle? rowStyle = row.StyleSnapshot ?? currentOpts.DefaultRowStyleSnapshot;
                    double rowGap = row.GapOverride ?? rowStyle?.Gap ?? PdfRowStyle.DefaultGap;
                double totalGap = rowGap * Math.Max(0, columns - 1);
                if (totalGap >= frameWidth) {
                    return rowStyle?.SpacingBefore ?? 0D;
                }

                double columnAreaWidth = frameWidth - totalGap;
                double tallestFirstVisual = 0D;
                for (int columnIndex = 0; columnIndex < columns; columnIndex++) {
                    RowColumn column = row.Columns[columnIndex];
                    if (column.Blocks.Count == 0) {
                        continue;
                    }

                    double columnWidth = Math.Max(0D, columnAreaWidth * (column.WidthPercent / 100D));
                    tallestFirstVisual = Math.Max(tallestFirstVisual, MeasureNextBlockFirstVisualHeight(column.Blocks[0], frameX, columnWidth, fontSize));
                }

                return (rowStyle?.SpacingBefore ?? 0D) + tallestFirstVisual;
            }

            return 0D;
        }

        void ConsumeSpacer(double height) {
            double remaining = height;
            while (remaining > 0.001D) {
                double available = y - currentOpts.MarginBottom;
                if (available <= 0.5D) {
                    NewPage();
                    continue;
                }

                double consumed = Math.Min(remaining, available);
                y -= consumed;
                remaining -= consumed;
                if (remaining > 0.001D) {
                    NewPage();
                }
            }
        }

        void ProcessBlocks(System.Collections.Generic.IEnumerable<IPdfBlock> sequence) {
            var blockList = sequence as System.Collections.Generic.IList<IPdfBlock> ?? sequence.ToList();
            for (int blockIndex = 0; blockIndex < blockList.Count; blockIndex++) {
                var block = blockList[blockIndex];
                IPdfBlock? nextBlock = blockIndex + 1 < blockList.Count ? blockList[blockIndex + 1] : null;
                if (block is PageBlock pageBlock) {
                    FlushPage(pageDirty || (currentPage?.Images.Count ?? 0) > 0 || (currentPage?.Annotations.Count ?? 0) > 0 || (currentPage?.GraphicsStates.Count ?? 0) > 0 || (currentPage?.Shadings.Count ?? 0) > 0 || (currentPage?.NamedDestinations.Count ?? 0) > 0);
                    optionsStack.Push(pageBlock.Options);
                    pageGroupStack.Push(currentPageGroupId);
                    currentOpts = pageBlock.Options;
                    currentPageGroupId = nextPageGroupId++;
                    currentPage = null;
                    StartPage(currentOpts);
                    ProcessBlocks(pageBlock.Blocks);
                    FlushPage(force: true);
                    optionsStack.Pop();
                    currentPageGroupId = pageGroupStack.Pop();
                    currentOpts = optionsStack.Peek();
                    currentPage = null;
                    continue;
                }

                EnsurePage();

                if (block is PageBreakBlock) { NewPage(); continue; }
                if (block is BookmarkBlock bookmark) { AddNamedDestination(bookmark, y); continue; }
                if (block is SpacerBlock spacer) { ConsumeSpacer(spacer.Height); continue; }
                if (block is HeadingBlock hb) {
                    PdfHeadingStyle? headingStyle = ResolveHeadingStyle(hb, currentOpts);
                    double size = GetHeadingFontSize(hb, headingStyle);
                    double leading = GetHeadingLeading(headingStyle, size);
                    double spacingBefore = headingStyle?.SpacingBefore ?? 0D;
                    double spacingAfter = GetHeadingSpacingAfter(headingStyle, leading);
                    var headingFont = ChooseBold(ChooseNormal(currentOpts.DefaultFont));
                    var lines = WrapSimpleText(hb.Text, width, headingFont, size);
                    double needed = spacingBefore + lines.Count * leading + spacingAfter;
                    bool keepWithNext = headingStyle?.KeepWithNext ?? true;
                    if (keepWithNext && nextBlock != null) {
                        double keepHeight = needed + MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (keepHeight > needed + 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    if (y - needed < currentOpts.MarginBottom) { NewPage(); }
                    if (spacingBefore > 0) {
                        y -= spacingBefore;
                    }

                    if (currentOpts.CreateOutlineFromHeadings) {
                        currentPage!.Bookmarks.Add(new PageBookmark { Level = hb.Level, Title = hb.Text, Y = y });
                    }
                    double firstBaseline = FirstTextBaselineFromTop(headingFont, size, y);
                    AddHeadingLinkAnnotations(hb, lines, headingFont, size, leading, currentOpts.MarginLeft, width, firstBaseline);
                    WriteLines("F2", size, leading, currentOpts.MarginLeft, firstBaseline, lines, hb.Align, hb.Color ?? headingStyle?.Color, applyBaselineTweak: false);
                    currentPage!.UsedBold = true;
                    usedBold = true;
                    y -= lines.Count * leading + spacingAfter;
                } else if (block is RichParagraphBlock rpb) {
                    double size = currentOpts.DefaultFontSize;
                    PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(rpb);
                    double leading = GetParagraphLeading(paragraphStyle, size);
                    double spacingBefore = GetParagraphSpacingBefore(paragraphStyle);
                    double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
                    var textFrame = GetParagraphTextFrame(paragraphStyle, currentOpts.MarginLeft, width);
                    var (lines, lineHeights) = WrapRichRuns(rpb.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, GetParagraphTabStopWidth(paragraphStyle));
                    if (paragraphStyle?.KeepWithNext == true && nextBlock != null && lines.Count > 0) {
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                        double keepHeight = spacingBefore + lineHeights.Sum() + spacingAfter + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    if (paragraphStyle?.KeepTogether == true) {
                        double paragraphHeight = spacingBefore + lineHeights.Sum();
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (paragraphHeight > availableHeight + 0.001) {
                            throw new ArgumentException("Paragraph height exceeds the available page content height.");
                        }

                        if (y < yStart - 0.001 && y - paragraphHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    int lineIndex = 0;
                    bool firstSegment = true;
                    while (lineIndex < lines.Count) {
                        double available = y - currentOpts.MarginBottom;
                        if (available <= 0.5) {
                            NewPage();
                            firstSegment = false;
                            continue;
                        }

                        double segmentSpacingBefore = firstSegment && y < yStart - 0.001 ? spacingBefore : 0;
                        double minimumLineHeight = lineHeights[lineIndex];
                        if (available < segmentSpacingBefore + minimumLineHeight) {
                            NewPage();
                            available = y - currentOpts.MarginBottom;
                            if (y >= yStart - 0.001) {
                                segmentSpacingBefore = 0;
                            }
                            if (available < segmentSpacingBefore + minimumLineHeight) {
                                segmentSpacingBefore = Math.Max(0, available - minimumLineHeight);
                            }
                        }

                        double roomForText = Math.Max(0, available - segmentSpacingBefore);
                        int take = 0;
                        double heightSum = 0;
                        for (int k = lineIndex; k < lines.Count; k++) {
                            double lineHeight = lineHeights[k];
                            if (heightSum + lineHeight > roomForText) {
                                break;
                            }

                            heightSum += lineHeight;
                            take++;
                        }

                        if (TryApplyWidowControl(paragraphStyle, lines.Count, lineIndex, ref take, ref heightSum, lineHeights, y < yStart - 0.001)) {
                            NewPage();
                            firstSegment = false;
                            continue;
                        }

                        if (take == 0) {
                            NewPage();
                            firstSegment = false;
                            continue;
                        }

                        if (segmentSpacingBefore > 0) {
                            y -= segmentSpacingBefore;
                        }

                        var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                        var sliceHeights = new System.Collections.Generic.List<double>();
                        for (int k = 0; k < take; k++) {
                            sliceLines.Add(lines[lineIndex + k]);
                            sliceHeights.Add(lineHeights[lineIndex + k]);
                        }

                        bool sliceStartsAtFirstLine = lineIndex == 0;
                        pageDirty = true;
                        var paragraphFont = ChooseNormal(currentOpts.DefaultFont);
                        WriteRichParagraph(sb, rpb, sliceLines, sliceHeights, currentOpts, FirstTextBaselineFromTop(paragraphFont, size, y), size, leading, currentPage!.Annotations, textFrame.X, textFrame.Width, sliceStartsAtFirstLine ? textFrame.FirstLineX : null, sliceStartsAtFirstLine ? textFrame.FirstLineWidth : null);
                        y -= heightSum;
                        lineIndex += take;
                        firstSegment = false;
                        if (lineIndex < lines.Count) {
                            NewPage();
                        } else {
                            y -= spacingAfter;
                        }
                    }

                    if (rpb.Runs.Any(r => r.Bold)) { currentPage!.UsedBold = true; usedBold = true; }
                    if (rpb.Runs.Any(r => r.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
                    if (rpb.Runs.Any(r => r.Bold && r.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }
                } else if (block is BulletListBlock bl) {
                    PdfListStyle? listStyle = ResolveListStyle(bl, currentOpts);
                    double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
                    double leading = GetListLeading(listStyle, size);
                    var baseFont = ChooseNormal(currentOpts.DefaultFont);
                    const string bulletGlyph = "•";
                    double bulletWidth = EstimateSimpleTextWidth(bulletGlyph, baseFont, size);
                    double spaceAdvance = EstimateSimpleTextWidth(" ", baseFont, size);
                    double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
                    double indent = bulletWidth + markerGap;
                    double listLeftIndent = listStyle?.LeftIndent ?? 0D;
                    double rawTextWidth = width - listLeftIndent - indent;
                    double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidth("WW", baseFont, size));
                    double alignmentWidth = Math.Max(0, rawTextWidth);
                    double itemSpacing = GetListItemSpacing(listStyle, leading);
                    var wrappedItems = new System.Collections.Generic.List<System.Collections.Generic.List<string>>(bl.Items.Count);
                    for (int itemIndex = 0; itemIndex < bl.Items.Count; itemIndex++) {
                        wrappedItems.Add(WrapSimpleText(bl.Items[itemIndex], availableWidth, baseFont, size));
                    }

                    double listSpacingBefore = listStyle?.SpacingBefore ?? 0D;
                    double listSpacingAfter = listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing;
                    double listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                    if (listStyle?.KeepTogether == true) {
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (listHeight > availableHeight + 0.001) {
                            throw new ArgumentException("List height exceeds the available page content height.");
                        }

                        if (y < yStart - 0.001 && y - listHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    if (listStyle?.KeepWithNext == true && nextBlock != null && wrappedItems.Count > 0) {
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                        double keepHeight = listHeight + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    for (int itemIndex = 0; itemIndex < bl.Items.Count; itemIndex++) {
                        var lines = wrappedItems[itemIndex];
                        double firstLineWidth = lines.Count > 0 ? EstimateSimpleTextWidth(lines[0], baseFont, size) : 0;
                        double firstLineDx = 0;
                        if (bl.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                        else if (bl.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);

                        double spacingBefore = itemIndex == 0 ? listSpacingBefore : 0D;
                        double spacingAfter = itemIndex == bl.Items.Count - 1 ? listSpacingAfter : itemSpacing;
                        RenderListItem(lines, bulletGlyph, currentOpts.MarginLeft + listLeftIndent + firstLineDx, bulletWidth, PdfAlign.Left, currentOpts.MarginLeft + listLeftIndent + indent, alignmentWidth, bl.Align, bl.Color ?? listStyle?.Color, size, leading, spacingBefore, spacingAfter);
                    }
                } else if (block is NumberedListBlock nl) {
                    PdfListStyle? listStyle = ResolveListStyle(nl, currentOpts);
                    double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
                    double leading = GetListLeading(listStyle, size);
                    var baseFont = ChooseNormal(currentOpts.DefaultFont);
                    int lastNumber = nl.StartNumber + Math.Max(0, nl.Items.Count - 1);
                    string widestMarker = lastNumber.ToString(CultureInfo.InvariantCulture) + ".";
                    double markerWidth = EstimateSimpleTextWidth(widestMarker, baseFont, size);
                    double spaceAdvance = EstimateSimpleTextWidth(" ", baseFont, size);
                    double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
                    double indent = markerWidth + markerGap;
                    double rawTextWidth = width - (listStyle?.LeftIndent ?? 0D) - indent;
                    double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidth("WW", baseFont, size));
                    double alignmentWidth = Math.Max(0, rawTextWidth);
                    double itemSpacing = GetListItemSpacing(listStyle, leading);
                    double listLeftIndent = listStyle?.LeftIndent ?? 0D;
                    var wrappedItems = new System.Collections.Generic.List<System.Collections.Generic.List<string>>(nl.Items.Count);
                    for (int itemIndex = 0; itemIndex < nl.Items.Count; itemIndex++) {
                        wrappedItems.Add(WrapSimpleText(nl.Items[itemIndex], availableWidth, baseFont, size));
                    }

                    double listSpacingBefore = listStyle?.SpacingBefore ?? 0D;
                    double listSpacingAfter = listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing;
                    double listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                    if (listStyle?.KeepTogether == true) {
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (listHeight > availableHeight + 0.001) {
                            throw new ArgumentException("List height exceeds the available page content height.");
                        }

                        if (y < yStart - 0.001 && y - listHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    if (listStyle?.KeepWithNext == true && nextBlock != null && wrappedItems.Count > 0) {
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                        double keepHeight = listHeight + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    for (int itemIndex = 0; itemIndex < nl.Items.Count; itemIndex++) {
                        string marker = (nl.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + ".";
                        var lines = wrappedItems[itemIndex];
                        double firstLineWidth = lines.Count > 0 ? EstimateSimpleTextWidth(lines[0], baseFont, size) : 0;
                        double firstLineDx = 0;
                        if (nl.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                        else if (nl.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);

                        double spacingBefore = itemIndex == 0 ? listSpacingBefore : 0D;
                        double spacingAfter = itemIndex == nl.Items.Count - 1 ? listSpacingAfter : itemSpacing;
                        RenderListItem(lines, marker, currentOpts.MarginLeft + listLeftIndent + firstLineDx, markerWidth, PdfAlign.Right, currentOpts.MarginLeft + listLeftIndent + indent, alignmentWidth, nl.Align, nl.Color ?? listStyle?.Color, size, leading, spacingBefore, spacingAfter);
                    }
                } else if (block is TableBlock tb) {
                    var style = tb.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
                    int cols = GetTableColumnCount(tb);
                    if (cols == 0) continue;
                    double padLeft = GetTableCellPaddingLeft(style);
                    double padRight = GetTableCellPaddingRight(style);
                    double padTop = GetTableCellPaddingTop(style);
                    double padBottom = GetTableCellPaddingBottom(style);
                    double colGapPx = 0;
                    double size = GetTableBodyFontSize(style, currentOpts.DefaultFontSize);
                    if (!IsValidPdfAlign(tb.Align)) {
                        throw new ArgumentException("Table alignment must be Left, Center, or Right.");
                    }
                    if (style.Alignments != null) {
                        foreach (var alignment in style.Alignments) {
                            if (!IsValidPdfColumnAlign(alignment)) {
                                throw new ArgumentException("Table column alignments must be Left, Center, or Right.");
                            }
                        }
                    }
                    if (style.VerticalAlignments != null) {
                        foreach (var verticalAlignment in style.VerticalAlignments) {
                            if (!IsValidPdfCellVerticalAlign(verticalAlignment)) {
                                throw new ArgumentException("Table vertical alignments must be defined PDF cell vertical alignment values.");
                            }
                        }
                    }
                    if (!IsValidPdfAlign(style.CaptionAlign)) {
                        throw new ArgumentException("Table caption alignment must be Left, Center, or Right.");
                    }
                    if (style.BorderWidth < 0 || double.IsNaN(style.BorderWidth) || double.IsInfinity(style.BorderWidth)) {
                        throw new ArgumentException("Table border width must be a non-negative finite value.");
                    }
                    if (style.RowSeparatorWidth < 0 || double.IsNaN(style.RowSeparatorWidth) || double.IsInfinity(style.RowSeparatorWidth)) {
                        throw new ArgumentException("Table row separator width must be a non-negative finite value.");
                    }
                    if (style.HeaderSeparatorWidth < 0 || double.IsNaN(style.HeaderSeparatorWidth) || double.IsInfinity(style.HeaderSeparatorWidth)) {
                        throw new ArgumentException("Table header separator width must be a non-negative finite value.");
                    }
                    if (style.FooterSeparatorWidth < 0 || double.IsNaN(style.FooterSeparatorWidth) || double.IsInfinity(style.FooterSeparatorWidth)) {
                        throw new ArgumentException("Table footer separator width must be a non-negative finite value.");
                    }
                    if (style.CellPaddingX < 0 || double.IsNaN(style.CellPaddingX) || double.IsInfinity(style.CellPaddingX)) {
                        throw new ArgumentException("Table horizontal cell padding must be a non-negative finite value.");
                    }
                    if (style.CellPaddingY < 0 || double.IsNaN(style.CellPaddingY) || double.IsInfinity(style.CellPaddingY)) {
                        throw new ArgumentException("Table vertical cell padding must be a non-negative finite value.");
                    }
                    if (style.MinRowHeight < 0 || double.IsNaN(style.MinRowHeight) || double.IsInfinity(style.MinRowHeight)) {
                        throw new ArgumentException("Table minimum row height must be a non-negative finite value.");
                    }
                    if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
                        throw new ArgumentException("Table spacing before must be a non-negative finite value.");
                    }
                    if (style.Caption != null && string.IsNullOrWhiteSpace(style.Caption)) {
                        throw new ArgumentException("Table caption cannot be empty or whitespace.");
                    }
                    if (style.CaptionFontSize.HasValue && (style.CaptionFontSize.Value <= 0 || double.IsNaN(style.CaptionFontSize.Value) || double.IsInfinity(style.CaptionFontSize.Value))) {
                        throw new ArgumentException("Table caption font size must be a positive finite value.");
                    }
                    if (style.CaptionSpacingAfter < 0 || double.IsNaN(style.CaptionSpacingAfter) || double.IsInfinity(style.CaptionSpacingAfter)) {
                        throw new ArgumentException("Table caption spacing after must be a non-negative finite value.");
                    }
                    if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
                        throw new ArgumentException("Table spacing after must be a non-negative finite value.");
                    }
                    if (double.IsNaN(style.RowBaselineOffset) || double.IsInfinity(style.RowBaselineOffset)) {
                        throw new ArgumentException("Table row baseline offset must be a finite value.");
                    }
                    if (style.CellFills != null) {
                        foreach (var cellFill in style.CellFills) {
                            if (cellFill.Key.Row < 0 || cellFill.Key.Column < 0) {
                                throw new ArgumentException("Table cell fill coordinates cannot be negative.");
                            }
                        }
                    }
                    if (style.CellBorders != null) {
                        foreach (var cellBorder in style.CellBorders) {
                            if (cellBorder.Key.Row < 0 || cellBorder.Key.Column < 0) {
                                throw new ArgumentException("Table cell border coordinates cannot be negative.");
                            }
                            if (cellBorder.Value == null || cellBorder.Value.Width < 0 || double.IsNaN(cellBorder.Value.Width) || double.IsInfinity(cellBorder.Value.Width)) {
                                throw new ArgumentException("Table cell border widths must be non-negative finite values.");
                            }
                        }
                    }
                    if (style.HeaderRowCount < 0) {
                        throw new ArgumentException("Table header row count cannot be negative.");
                    }
                    if (style.FooterRowCount < 0) {
                        throw new ArgumentException("Table footer row count cannot be negative.");
                    }

                    ValidateTableRoleRowCounts(style, tb.Rows.Count);
                    int headerRowCount = style.HeaderRowCount;
                    int footerRowCount = style.FooterRowCount;
                    int footerStartRowIndex = tb.Rows.Count - footerRowCount;
                    ValidateTableCellStyleCoordinates(style, tb.Rows.Count, cols);
                    ValidateTableColumnStyleBounds(style, cols);
                    ValidateTableRowSpansWithinRoleBoundaries(tb, cols, headerRowCount, footerStartRowIndex);
                    double[]? autoFitWeights = style.AutoFitColumns
                        ? MeasureAutoFitColumnWeights(tb, currentOpts, style, size, headerRowCount, footerStartRowIndex)
                        : null;
                    double[]? autoFitMinimumWidths = style.AutoFitColumns
                        ? MeasureAutoFitColumnMinimumWidths(tb, currentOpts, style, size, headerRowCount, footerStartRowIndex)
                        : null;
                    double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                    double tableWidth = ResolveTableAvailableWidth(style, contentWidth);
                    double[] colPixel = new double[cols];
                    double[] colWeights = new double[cols];
                    bool[] fixedColumns = new bool[cols];
                    double?[] minWidths = new double?[cols];
                    double?[] maxWidths = new double?[cols];
                    double fixedWidthTotal = 0;
                    double totalWeight = 0;
                    for (int c = 0; c < cols; c++) {
                        double? minWidth = GetOptionalColumnWidth(style.ColumnMinWidthPoints, c, "Table minimum column widths must be positive finite values.");
                        if (!minWidth.HasValue && autoFitMinimumWidths != null && c < autoFitMinimumWidths.Length) {
                            minWidth = autoFitMinimumWidths[c];
                        }

                        double? maxWidth = GetOptionalColumnWidth(style.ColumnMaxWidthPoints, c, "Table maximum column widths must be positive finite values.");
                        if (minWidth.HasValue && maxWidth.HasValue && minWidth.Value > maxWidth.Value + 0.001) {
                            throw new ArgumentException("Table minimum column widths cannot be greater than maximum column widths.");
                        }

                        minWidths[c] = minWidth;
                        maxWidths[c] = maxWidth;

                        if (style.ColumnWidthPoints != null &&
                            c < style.ColumnWidthPoints.Count &&
                            style.ColumnWidthPoints[c].HasValue) {
                            double fixedWidth = style.ColumnWidthPoints[c]!.Value;
                            if (fixedWidth <= 0 || double.IsNaN(fixedWidth) || double.IsInfinity(fixedWidth)) {
                                throw new ArgumentException("Table fixed column widths must be positive finite values.");
                            }
                            if (minWidth.HasValue && fixedWidth < minWidth.Value - 0.001) {
                                throw new ArgumentException("Table fixed column widths cannot be smaller than configured minimum widths.");
                            }
                            if (maxWidth.HasValue && fixedWidth > maxWidth.Value + 0.001) {
                                throw new ArgumentException("Table fixed column widths cannot be greater than configured maximum widths.");
                            }

                            colPixel[c] = fixedWidth;
                            fixedColumns[c] = true;
                            fixedWidthTotal += fixedWidth;
                            continue;
                        }

                        double weight = 1;
                        if (style.ColumnWidthWeights != null && c < style.ColumnWidthWeights.Count) {
                            weight = style.ColumnWidthWeights[c];
                            if (weight <= 0 || double.IsNaN(weight) || double.IsInfinity(weight)) {
                                throw new ArgumentException("Table column width weights must be positive finite values.");
                            }
                        } else if (autoFitWeights != null && c < autoFitWeights.Length) {
                            weight = autoFitWeights[c];
                        }

                        colWeights[c] = weight;
                        totalWeight += weight;
                    }
                    double tableInnerWidth = tableWidth - (cols - 1) * colGapPx;
                    if (fixedWidthTotal > tableInnerWidth + 0.001) {
                        throw new ArgumentException("Table fixed column widths exceed the available table width.");
                    }

                    if (totalWeight <= 0) {
                        tableInnerWidth = fixedWidthTotal;
                        tableWidth = tableInnerWidth + (cols - 1) * colGapPx;
                    }

                    double remainingWidth = Math.Max(0, tableInnerWidth - fixedWidthTotal);
                    DistributeFlexibleColumns(colPixel, colWeights, fixedColumns, minWidths, maxWidths, remainingWidth);
                    double usedTableInnerWidth = colPixel.Sum();
                    if (usedTableInnerWidth < tableInnerWidth - 0.001) {
                        tableInnerWidth = usedTableInnerWidth;
                        tableWidth = tableInnerWidth + (cols - 1) * colGapPx;
                    }
                    for (int c = 0; c < cols; c++) {
                        if (colPixel[c] - padLeft - padRight <= 0.001) {
                            throw new ArgumentException("Table horizontal cell padding must leave a positive text width.");
                        }
                    }

                    var rowLines = new string[tb.Rows.Count][][];
                    var rowLineCounts = new int[tb.Rows.Count];
                    var rowHeights = new double[tb.Rows.Count];
                    var rowLeadings = new double[tb.Rows.Count];
                    var rowSizes = new double[tb.Rows.Count];
                    var rowBold = new bool[tb.Rows.Count];
                    for (int ri = 0; ri < tb.Rows.Count; ri++) {
                        double rowSize = GetTableRowFontSize(style, ri, headerRowCount, footerStartRowIndex, currentOpts.DefaultFontSize);
                        double rowLeading = GetTableLeading(style, rowSize);
                        bool rowUsesBold = GetTableRowBold(style, ri, headerRowCount, footerStartRowIndex);
                        rowSizes[ri] = rowSize;
                        rowLeadings[ri] = rowLeading;
                        rowBold[ri] = rowUsesBold;
                        rowLines[ri] = new string[cols][];
                        int maxLines = 1;
                        for (int ci = 0; ci < cols; ci++) {
                            rowLines[ri][ci] = System.Array.Empty<string>();
                        }

                        var cells = GetTableCellLayouts(tb, ri, cols);
                        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                            TableCellLayout cell = cells[cellIndex];
                            var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                            double cellWidth = GetTableCellWidth(colPixel, cell.Column, cell.ColumnSpan, colGapPx);
                            double innerWidth = Math.Max(1, cellWidth - padLeft - padRight);
                            var lines = WrapSimpleText(cell.Text, innerWidth, cellFont, rowSize);
                            rowLines[ri][cell.Column] = lines.ToArray();
                            if (cell.RowSpan <= 1) {
                                maxLines = Math.Max(maxLines, lines.Count);
                            }
                        }
                        rowLineCounts[ri] = maxLines;
                        rowHeights[ri] = Math.Max(maxLines * rowLeading + padTop + padBottom, style.MinRowHeight);
                    }
                    ApplyTableRowSpanHeights(tb, cols, rowLines, rowHeights, rowLeadings, padTop, padBottom);
                    double xOrigin = ResolveTableX(tb.Align, style, currentOpts.MarginLeft, contentWidth, tableWidth);

                    double maxContentHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                    string? captionText = string.IsNullOrWhiteSpace(style.Caption) ? null : style.Caption;
                    System.Collections.Generic.List<string>? captionLines = null;
                    double captionSize = style.CaptionFontSize ?? size;
                    double captionLeading = captionSize * 1.25;
                    double captionHeight = 0;
                    if (captionText != null) {
                        var captionFontForWrap = ChooseNormal(currentOpts.DefaultFont);
                        captionLines = WrapSimpleText(captionText, tableWidth, captionFontForWrap, captionSize).ToList();
                        captionHeight = captionLines.Count * captionLeading;
                        double firstRowHeight = rowHeights.Length > 0 ? rowHeights[0] : 0;
                        if (captionHeight + style.CaptionSpacingAfter + firstRowHeight > maxContentHeight + 0.001) {
                            throw new ArgumentException("Table caption and first row exceed the available page content height.");
                        }
                    }

                    double tableContentHeight = (captionLines == null ? 0 : captionHeight + style.CaptionSpacingAfter) + rowHeights.Sum();
                    if (style.KeepTogether) {
                        double keepHeight = style.SpacingBefore + tableContentHeight + style.SpacingAfter;
                        if (keepHeight > maxContentHeight + 0.001) {
                            throw new ArgumentException("Table height exceeds the available page content height.");
                        }

                        if (y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    if (style.KeepWithNext && nextBlock != null) {
                        double tableHeight = style.SpacingBefore + tableContentHeight + style.SpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        double keepHeight = tableHeight + nextHeight;
                        if (nextHeight > 0.001 && tableHeight <= maxContentHeight + 0.001 && keepHeight <= maxContentHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    if (style.SpacingBefore > 0) {
                        if (y < yStart - 0.001 && y - style.SpacingBefore < currentOpts.MarginBottom) {
                            NewPage();
                        }

                        y -= style.SpacingBefore;
                    }

                    if (captionLines != null) {
                        var captionFont = ChooseNormal(currentOpts.DefaultFont);
                        double firstRowHeight = rowHeights.Length > 0 ? rowHeights[0] : 0;
                        double captionAndFirstRowHeight = captionHeight + style.CaptionSpacingAfter + firstRowHeight;
                        if (y < yStart - 0.001 &&
                            y - Math.Min(captionAndFirstRowHeight, maxContentHeight) < currentOpts.MarginBottom) {
                            NewPage();
                        }

                        WriteLinesInternal("F1", captionSize, captionLeading, xOrigin, tableWidth, y - GetAscender(captionFont, captionSize), captionLines, style.CaptionAlign, style.CaptionColor);
                        y -= captionHeight + style.CaptionSpacingAfter;
                    }

                    if (TableUsesBold(style, tb.Rows.Count, headerRowCount, footerStartRowIndex)) {
                        currentPage!.UsedBold = true;
                        usedBold = true;
                    }

                    bool hasRepeatableHeader = headerRowCount > 0 && tb.Rows.Count > headerRowCount;
                    double repeatHeaderHeight = 0;
                    for (int i = 0; i < headerRowCount; i++) {
                        repeatHeaderHeight += rowHeights[i];
                    }

                    bool ShouldBreakBefore(double rowHeight) =>
                        y < yStart - 0.001 &&
                        y - rowHeight < currentOpts.MarginBottom &&
                        rowHeight <= maxContentHeight;

                    bool CanRepeatHeaderWithSegment(int rowIndex) =>
                        hasRepeatableHeader &&
                        rowIndex >= headerRowCount &&
                        repeatHeaderHeight + rowLeadings[rowIndex] + padTop + padBottom <= maxContentHeight + 0.001;

                    void DrawRepeatHeaders() {
                        for (int headerIndex = 0; headerIndex < headerRowCount; headerIndex++) {
                            DrawTableRow(headerIndex, renderAsHeader: true);
                        }
                    }

                    void NewTablePage(int rowIndex) {
                        NewPage();
                        if (CanRepeatHeaderWithSegment(rowIndex)) {
                            DrawRepeatHeaders();
                        }
                    }

                    void DrawTableRowSegment(int rowIndex, bool renderAsHeader, int startLine, int lineCount) {
                        bool renderAsFooter = rowIndex >= footerStartRowIndex;
                        bool rowUsesBold = rowBold[rowIndex];
                        double rowSize = rowSizes[rowIndex];
                        double rowLeading = rowLeadings[rowIndex];
                        if (rowUsesBold) {
                            currentPage!.UsedBold = true;
                            usedBold = true;
                        }

                        var cells = GetTableCellLayouts(tb, rowIndex, cols);
                        bool wholeRowSegment = startLine == 0 && lineCount == rowLineCounts[rowIndex];
                        double rowHeight = wholeRowSegment ? rowHeights[rowIndex] : Math.Max(1, lineCount) * rowLeading + padTop + padBottom;
                        double rowBottom = y - rowHeight;
                        if (currentOpts.Debug?.ShowTableRowBoxes == true) { pageDirty = true; DrawRowRect(sb, new PdfColor(1, 0, 1), 0.6, xOrigin, rowBottom, tableWidth, rowHeight); }
                        int bodyRowIndex = rowIndex - headerRowCount;
                        bool stripeBodyRow = bodyRowIndex >= 0 && bodyRowIndex % 2 == 1;
                        bool[] rowFillSkips = GetRowSpanContinuationSkipColumns(tb, rowIndex, cols);
                        if (style?.HeaderFill is not null && renderAsHeader) { pageDirty = true; DrawTableRowFill(sb, style.HeaderFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips); } else if (style?.FooterFill is not null && renderAsFooter) { pageDirty = true; DrawTableRowFill(sb, style.FooterFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips); } else if (!renderAsHeader && !renderAsFooter && style?.RowStripeFill is not null && stripeBodyRow) { pageDirty = true; DrawTableRowFill(sb, style.RowStripeFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips); }
                        if (!renderAsHeader && !renderAsFooter && style?.BodyColumnFills != null) {
                            bool[] bodyColumnFillSkips = GetMergedCellContinuationSkipColumns(tb, rowIndex, cols);
                            double fillX = xOrigin;
                            for (int fillColumn = 0; fillColumn < cols; fillColumn++) {
                                PdfColor? fill = fillColumn < style.BodyColumnFills.Count ? style.BodyColumnFills[fillColumn] : null;
                                if (fill.HasValue && (fillColumn >= bodyColumnFillSkips.Length || !bodyColumnFillSkips[fillColumn])) {
                                    pageDirty = true;
                                    DrawRowFill(sb, fill.Value, fillX, rowBottom, colPixel[fillColumn], rowHeight);
                                }
                                fillX += colPixel[fillColumn] + colGapPx;
                            }
                        }
                        if (style?.CellFills != null && style.CellFills.Count > 0) {
                            double fillX = xOrigin;
                            for (int fillColumn = 0; fillColumn < cols; fillColumn++) {
                                if (style.CellFills.TryGetValue((rowIndex, fillColumn), out PdfColor fill) &&
                                    TryGetTableCellLayoutAtColumn(cells, fillColumn, out TableCellLayout fillCell) &&
                                    (fillColumn >= rowFillSkips.Length || !rowFillSkips[fillColumn])) {
                                    pageDirty = true;
                                    int span = wholeRowSegment ? fillCell.ColumnSpan : 1;
                                    double fillHeight = rowHeight;
                                    double fillBottom = rowBottom;
                                    if (wholeRowSegment) {
                                        if (fillCell.RowSpan > 1) {
                                            fillHeight = GetTableCellHeight(rowHeights, rowIndex, fillCell.RowSpan);
                                            fillBottom = y - fillHeight;
                                        }
                                    }

                                    DrawRowFill(sb, fill, fillX, fillBottom, GetTableCellWidth(colPixel, fillColumn, span, colGapPx), fillHeight);
                                }
                                fillX += colPixel[fillColumn] + colGapPx;
                            }
                        }
                        if (currentOpts.Debug?.ShowTableBaselines == true) {
                            double x1 = xOrigin;
                            double x2 = xOrigin + tableWidth;
                            double baselineYDbg = y - padTop - GetAscender(GetTableRowFont(currentOpts, rowUsesBold), rowSize);
                            pageDirty = true;
                            DrawHLine(sb, new PdfColor(0, 0.6, 0), 0.4, x1, x2, baselineYDbg);
                        }
                        double xi = xOrigin;
                        double yRect = rowBottom;
                        double rowWidth = tableWidth;
                        double hRect = rowHeight;
                        string fontRes = GetTableRowFontResource(rowUsesBold);
                        var textColor = renderAsHeader ? style?.HeaderTextColor : renderAsFooter ? style?.FooterTextColor : style?.TextColor;
                        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                            TableCellLayout cell = cells[cellIndex];
                            int c = cell.Column;
                            xi = xOrigin;
                            for (int xColumn = 0; xColumn < c; xColumn++) {
                                xi += colPixel[xColumn] + colGapPx;
                            }

                            double cellWidth = GetTableCellWidth(colPixel, c, cell.ColumnSpan, colGapPx);
                            double innerW = cellWidth - padLeft - padRight;
                            double cellHeight = wholeRowSegment && cell.RowSpan > 1 ? GetTableCellHeight(rowHeights, rowIndex, cell.RowSpan) : rowHeight;
                            double cellBottom = y - cellHeight;
                            var align = PdfColumnAlign.Left;
                            if (style?.Alignments != null && c < style.Alignments.Count) align = style.Alignments[c];
                            if (style?.RightAlignNumeric == true && LooksNumeric(cell.Text)) align = PdfColumnAlign.Right;
                            var verticalAlign = PdfCellVerticalAlign.Top;
                            if (style?.VerticalAlignments != null && c < style.VerticalAlignments.Count) verticalAlign = style.VerticalAlignments[c];

                            var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                            var lines = rowLines[rowIndex][c];
                            int sourceStartLine = wholeRowSegment && cell.RowSpan > 1 ? 0 : startLine;
                            int requestedLineCount = wholeRowSegment && cell.RowSpan > 1 ? lines.Length : lineCount;
                            int visibleLineCount = Math.Max(0, Math.Min(requestedLineCount, lines.Length - sourceStartLine));
                            double verticalOffset = 0;
                            if (visibleLineCount > 0) {
                                double availableTextHeight = Math.Max(0, cellHeight - padTop - padBottom);
                                double visibleTextHeight = visibleLineCount * rowLeading;
                                double unusedTextHeight = Math.Max(0, availableTextHeight - visibleTextHeight);
                                if (verticalAlign == PdfCellVerticalAlign.Middle) verticalOffset = unusedTextHeight / 2;
                                else if (verticalAlign == PdfCellVerticalAlign.Bottom) verticalOffset = unusedTextHeight;
                            }

                            double firstBaseline = y - padTop - verticalOffset - GetAscender(cellFont, rowSize) + (style?.RowBaselineOffset ?? 0);

                            pageDirty = true;
                            for (int lineIndex = 0; lineIndex < visibleLineCount; lineIndex++) {
                                int sourceLineIndex = sourceStartLine + lineIndex;
                                if (sourceLineIndex >= lines.Length) {
                                    continue;
                                }

                                string line = lines[sourceLineIndex];
                                double textW = EstimateSimpleTextWidth(line, cellFont, rowSize);
                                double offset = 0;
                                if (align == PdfColumnAlign.Center) offset = Math.Max(0, (innerW - textW) / 2);
                                else if (align == PdfColumnAlign.Right) offset = Math.Max(0, innerW - textW);
                                double xCell = xi + padLeft + offset;
                                double yCell = firstBaseline - lineIndex * rowLeading;
                                if (ShouldClipTableCellText(xCell, yCell, textW, cellFont, rowSize, xi, cellBottom, cellWidth, cellHeight)) {
                                    WriteClippedCell(sb, fontRes, rowSize, xCell, yCell, line, textColor, currentOpts, xi - TableCellClipBleed, cellBottom - TableCellClipBleed, cellWidth + TableCellClipBleed * 2D, cellHeight + TableCellClipBleed * 2D);
                                } else {
                                    WriteCell(sb, fontRes, rowSize, xCell, yCell, line, textColor, currentOpts);
                                }
                            }

                            string? linkUri = cell.LinkUri;
                            string? linkContents = cell.LinkContents;
                            if (tb.Links.TryGetValue((rowIndex, c), out var uri)) {
                                linkUri = uri;
                                linkContents = cell.Text;
                            }

                            if (!string.IsNullOrEmpty(linkUri)) {
                                double x1 = xi + padLeft;
                                double x2 = xi + padLeft + innerW;
                                double y1 = cellBottom + padBottom;
                                double y2 = y - padTop;
                                currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = linkUri!, Contents = linkContents ?? cell.Text });
                            }
                        }
                        if (style?.BorderColor is not null && style.BorderWidth > 0) {
                            pageDirty = true;
                            bool[] topBorderSkips = GetRowSpanBoundarySkipColumns(tb, rowIndex - 1, cols);
                            bool[] bottomBorderSkips = GetRowSpanBoundarySkipColumns(tb, rowIndex, cols);
                            bool segmentBorderRows = HasSkippedColumns(topBorderSkips, cols) || HasSkippedColumns(bottomBorderSkips, cols);
                            if (segmentBorderRows) {
                                DrawTableHorizontalLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, colPixel, colGapPx, yRect + hRect, topBorderSkips);
                                DrawTableHorizontalLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, colPixel, colGapPx, yRect, bottomBorderSkips);
                                DrawVLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, yRect + hRect, yRect);
                                DrawVLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin + tableWidth, yRect + hRect, yRect);
                            } else {
                                DrawRowRect(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, yRect, rowWidth, hRect);
                            }

                            double xi2 = xOrigin;
                            double yTop = yRect + hRect;
                            double yBottom = yRect;
                            for (int c = 0; c < cols - 1; c++) {
                                xi2 += colPixel[c];
                                if (IsTableBoundaryInsideSpannedCell(tb, rowIndex, c, cols)) {
                                    xi2 += colGapPx;
                                    continue;
                                }

                                if (currentOpts.Debug?.ShowTableColumnGuides == true)
                                    DrawVLine(sb, new PdfColor(0, 0, 1), Math.Max(0.3, style.BorderWidth), xi2, yTop, yBottom);
                                else
                                    DrawVLine(sb, style.BorderColor.Value, style.BorderWidth, xi2, yTop, yBottom);
                                xi2 += colGapPx;
                            }
                        }
                        if (style != null && renderAsFooter && rowIndex == footerStartRowIndex) {
                            PdfColor? footerSeparatorColor = style.FooterSeparatorColor ?? style.RowSeparatorColor;
                            double footerSeparatorWidth = style.FooterSeparatorWidth > 0 ? style.FooterSeparatorWidth : style.RowSeparatorWidth;
                            if (footerSeparatorColor is not null && footerSeparatorWidth > 0) {
                                pageDirty = true;
                                DrawTableHorizontalLine(sb, footerSeparatorColor.Value, footerSeparatorWidth, xOrigin, colPixel, colGapPx, y, GetRowSpanBoundarySkipColumns(tb, rowIndex - 1, cols));
                            }
                        }
                        PdfColor? separatorColor = renderAsHeader && style?.HeaderSeparatorColor is not null ? style.HeaderSeparatorColor : style?.RowSeparatorColor;
                        double separatorWidth = renderAsHeader && style?.HeaderSeparatorWidth > 0 ? style.HeaderSeparatorWidth : style?.RowSeparatorWidth ?? 0;
                        if (separatorColor is not null && separatorWidth > 0) {
                            pageDirty = true;
                            DrawTableHorizontalLine(sb, separatorColor.Value, separatorWidth, xOrigin, colPixel, colGapPx, rowBottom, GetRowSpanBoundarySkipColumns(tb, rowIndex, cols));
                        }
                        if (style?.CellBorders != null && style.CellBorders.Count > 0) {
                            double borderX = xOrigin;
                            for (int borderColumn = 0; borderColumn < cols; borderColumn++) {
                                if (style.CellBorders.TryGetValue((rowIndex, borderColumn), out PdfCellBorder? cellBorder) &&
                                    TryGetTableCellLayoutAtColumn(cells, borderColumn, out TableCellLayout borderCell) &&
                                    (borderColumn >= rowFillSkips.Length || !rowFillSkips[borderColumn]) &&
                                    cellBorder != null &&
                                    cellBorder.Color.HasValue &&
                                    cellBorder.Width > 0) {
                                    int span = wholeRowSegment ? borderCell.ColumnSpan : 1;
                                    double borderHeight = hRect;
                                    double borderBottom = yRect;
                                    if (wholeRowSegment) {
                                        if (borderCell.RowSpan > 1) {
                                            borderHeight = GetTableCellHeight(rowHeights, rowIndex, borderCell.RowSpan);
                                            borderBottom = y - borderHeight;
                                        }
                                    }

                                    pageDirty = true;
                                    DrawCellBorder(sb, cellBorder, borderX, borderBottom, GetTableCellWidth(colPixel, borderColumn, span, colGapPx), borderHeight);
                                }
                                borderX += colPixel[borderColumn] + colGapPx;
                            }
                        }
                        y -= rowHeight;
                    }

                    void DrawTableRow(int rowIndex, bool renderAsHeader) =>
                        DrawTableRowSegment(rowIndex, renderAsHeader, 0, rowLineCounts[rowIndex]);

                    void DrawSplitTableRow(int rowIndex, bool renderAsHeader) {
                        int startLine = 0;
                        int totalLines = rowLineCounts[rowIndex];
                        while (startLine < totalLines) {
                            double available = y - currentOpts.MarginBottom;
                            double minimumRowSegmentHeight = rowLeadings[rowIndex] + padTop + padBottom;
                            if (available < minimumRowSegmentHeight - 0.001) {
                                NewTablePage(rowIndex);
                                available = y - currentOpts.MarginBottom;
                            }

                            int maxLinesThisPage = Math.Max(1, (int)Math.Floor((available - padTop - padBottom) / rowLeadings[rowIndex]));
                            int take = Math.Min(totalLines - startLine, maxLinesThisPage);
                            DrawTableRowSegment(rowIndex, renderAsHeader && startLine == 0, startLine, take);
                            startLine += take;

                            if (startLine < totalLines) {
                                NewTablePage(rowIndex);
                            }
                        }
                    }

                    for (int rowIndex = 0; rowIndex < tb.Rows.Count; rowIndex++) {
                        if (rowHeights[rowIndex] > maxContentHeight + 0.001) {
                            if (!style.AllowRowBreakAcrossPages) {
                                throw new ArgumentException("Table row height exceeds the available page content height and row splitting is disabled.");
                            }

                            DrawSplitTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                            continue;
                        }

                        if (ShouldBreakBefore(rowHeights[rowIndex])) {
                            NewPage();
                            if (hasRepeatableHeader && rowIndex >= headerRowCount && repeatHeaderHeight + rowHeights[rowIndex] <= maxContentHeight + 0.001) {
                                DrawRepeatHeaders();
                            }
                        }

                        DrawTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                    }

                    y -= style.SpacingAfter;
                } else if (block is HorizontalRuleBlock hr) {
                    PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(hr, currentOpts);
                    ValidateHorizontalRule(ruleStyle);
                    if (ruleStyle.KeepWithNext && nextBlock != null) {
                        double needed = ruleStyle.SpacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        KeepFixedBlockWithNext(needed, nextHeight);
                    }

                    RenderHorizontalRuleBlock(hr, currentOpts.MarginLeft, width);
                } else if (block is ShapeBlock sbk) {
                    PdfDrawingStyle shapeStyle = ResolveDrawingStyle(sbk, currentOpts);
                    PdfDoc.ValidateDrawingStyle(shapeStyle, "Shape");
                    if (shapeStyle.KeepWithNext && nextBlock != null) {
                        double needed = shapeStyle.SpacingBefore + sbk.Shape.Height + shapeStyle.SpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        KeepFixedBlockWithNext(needed, nextHeight);
                    }

                    RenderShapeBlock(sbk, currentOpts.MarginLeft, width);
                } else if (block is DrawingBlock dbk) {
                    PdfDrawingStyle drawingStyle = ResolveDrawingStyle(dbk, currentOpts);
                    PdfDoc.ValidateDrawingStyle(drawingStyle, "Drawing");
                    if (drawingStyle.KeepWithNext && nextBlock != null) {
                        double needed = drawingStyle.SpacingBefore + dbk.Drawing.Height + drawingStyle.SpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        KeepFixedBlockWithNext(needed, nextHeight);
                    }

                    RenderDrawingBlock(dbk, currentOpts.MarginLeft, width);
                } else if (block is RowBlock rb) {
                    double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                    int ncols = rb.Columns.Count;
                    PdfRowStyle? rowStyle = rb.StyleSnapshot ?? currentOpts.DefaultRowStyleSnapshot;
                    double rowGap = rb.GapOverride ?? rowStyle?.Gap ?? PdfRowStyle.DefaultGap;
                    double rowSpacingBefore = rowStyle?.SpacingBefore ?? 0D;
                    double rowSpacingAfter = rowStyle?.SpacingAfter ?? 0D;
                    double totalGap = rowGap * Math.Max(0, ncols - 1);
                    if (totalGap >= contentWidth) {
                        throw new ArgumentException("Row column gaps must be smaller than the available page content width.");
                    }

                    double columnAreaWidth = contentWidth - totalGap;
                    double[] colXs = new double[ncols];
                    double[] colWs = new double[ncols];
                    double xAcc = currentOpts.MarginLeft;
                    for (int i = 0; i < ncols; i++) { double wCol = Math.Max(0, columnAreaWidth * (rb.Columns[i].WidthPercent / 100.0)); colXs[i] = xAcc; colWs[i] = wCol; xAcc += wCol + rowGap; }

                    var colStates = new System.Collections.Generic.List<(int idx, int line, int subline)>(ncols);
                    var colItems = new System.Collections.Generic.List<System.Collections.Generic.List<ColItem>>(ncols);
                    for (int i = 0; i < ncols; i++) {
                        colStates.Add((0, 0, 0));
                        var items = new System.Collections.Generic.List<ColItem>();
                        foreach (var cb in rb.Columns[i].Blocks) {
                            if (cb is HeadingBlock hb2) {
                                PdfHeadingStyle? headingStyle = ResolveHeadingStyle(hb2, currentOpts);
                                double size = GetHeadingFontSize(hb2, headingStyle);
                                double leading = GetHeadingLeading(headingStyle, size);
                                var headingFont = ChooseBold(ChooseNormal(currentOpts.DefaultFont));
                                var lines = WrapSimpleText(hb2.Text, colWs[i], headingFont, size);
                                items.Add(new ColHead {
                                    Block = hb2,
                                    Lines = lines,
                                    Leading = leading,
                                    Size = size,
                                    SpacingBefore = headingStyle?.SpacingBefore ?? 0D,
                                    SpacingAfter = GetHeadingSpacingAfter(headingStyle, leading),
                                    KeepWithNext = headingStyle?.KeepWithNext ?? true,
                                    Color = hb2.Color ?? headingStyle?.Color
                                });
                            } else if (cb is RichParagraphBlock rpb2) {
                                double size = currentOpts.DefaultFontSize;
                                PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(rpb2);
                                double leading = GetParagraphLeading(paragraphStyle, size);
                                var textFrame = GetParagraphTextFrame(paragraphStyle, 0, colWs[i]);
                                var wrap = WrapRichRuns(rpb2.Runs, textFrame.Width, size, ChooseNormal(currentOpts.DefaultFont), leading, textFrame.FirstLineWidth, GetParagraphTabStopWidth(paragraphStyle));
                                items.Add(new ColPar { Block = rpb2, Lines = wrap.Lines, Heights = wrap.LineHeights, Leading = leading, Size = size, XOffset = textFrame.X, TextWidth = textFrame.Width, FirstLineXOffset = textFrame.FirstLineX, FirstLineTextWidth = textFrame.FirstLineWidth });
                            } else if (cb is BulletListBlock bl2) {
                                PdfListStyle? listStyle = ResolveListStyle(bl2, currentOpts);
                                double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
                                double leading = GetListLeading(listStyle, size);
                                var baseFont = ChooseNormal(currentOpts.DefaultFont);
                                const string bulletGlyph = "•";
                                double bulletWidth = EstimateSimpleTextWidth(bulletGlyph, baseFont, size);
                                double spaceAdvance = EstimateSimpleTextWidth(" ", baseFont, size);
                                double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
                                double indent = bulletWidth + markerGap;
                                double listLeftIndent = listStyle?.LeftIndent ?? 0D;
                                double rawTextWidth = colWs[i] - listLeftIndent - indent;
                                double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidth("WW", baseFont, size));
                                double alignmentWidth = Math.Max(0, rawTextWidth);
                                double itemSpacing = GetListItemSpacing(listStyle, leading);
                                var listItems = new System.Collections.Generic.List<ColListItem>(bl2.Items.Count);
                                for (int itemIndex = 0; itemIndex < bl2.Items.Count; itemIndex++) {
                                    string text = bl2.Items[itemIndex];
                                    var lines = WrapSimpleText(text, availableWidth, baseFont, size);
                                    double firstLineWidth = lines.Count > 0 ? EstimateSimpleTextWidth(lines[0], baseFont, size) : 0;
                                    double firstLineDx = 0;
                                    if (bl2.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                                    else if (bl2.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);
                                    double spacingBefore = itemIndex == 0 ? listStyle?.SpacingBefore ?? 0D : 0D;
                                    double spacingAfter = itemIndex == bl2.Items.Count - 1 ? listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing : itemSpacing;
                                    listItems.Add(new ColListItem { Lines = lines, Marker = bulletGlyph, MarkerXOffset = listLeftIndent + firstLineDx, MarkerWidth = bulletWidth, MarkerAlign = PdfAlign.Left, TextXOffset = listLeftIndent + indent, TextWidth = alignmentWidth, TextAlign = bl2.Align, Color = bl2.Color ?? listStyle?.Color, Leading = leading, Size = size, SpacingBefore = spacingBefore, SpacingAfter = spacingAfter });
                                }

                                if ((listStyle?.KeepTogether == true || listStyle?.KeepWithNext == true) && listItems.Count > 0) {
                                    double listGroupHeight = 0D;
                                    foreach (var listItem in listItems) {
                                        listGroupHeight += listItem.SpacingBefore + listItem.Lines.Count * listItem.Leading + listItem.SpacingAfter;
                                    }

                                    if (listStyle?.KeepTogether == true) {
                                        listItems[0].IsFirstInKeepGroup = true;
                                        foreach (var listItem in listItems) {
                                            listItem.KeepTogether = true;
                                            listItem.KeepGroupHeight = listGroupHeight;
                                        }
                                    }

                                    if (listStyle?.KeepWithNext == true) {
                                        listItems[0].IsFirstInKeepWithNextGroup = true;
                                        foreach (var listItem in listItems) {
                                            listItem.KeepWithNext = true;
                                            listItem.KeepWithNextGroupItemCount = listItems.Count;
                                            listItem.KeepWithNextGroupHeight = listGroupHeight;
                                        }
                                    }
                                }

                                items.AddRange(listItems);
                            } else if (cb is NumberedListBlock nl2) {
                                PdfListStyle? listStyle = ResolveListStyle(nl2, currentOpts);
                                double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
                                double leading = GetListLeading(listStyle, size);
                                var baseFont = ChooseNormal(currentOpts.DefaultFont);
                                int lastNumber = nl2.StartNumber + Math.Max(0, nl2.Items.Count - 1);
                                string widestMarker = lastNumber.ToString(CultureInfo.InvariantCulture) + ".";
                                double markerWidth = EstimateSimpleTextWidth(widestMarker, baseFont, size);
                                double spaceAdvance = EstimateSimpleTextWidth(" ", baseFont, size);
                                double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
                                double indent = markerWidth + markerGap;
                                double listLeftIndent = listStyle?.LeftIndent ?? 0D;
                                double rawTextWidth = colWs[i] - listLeftIndent - indent;
                                double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidth("WW", baseFont, size));
                                double alignmentWidth = Math.Max(0, rawTextWidth);
                                double itemSpacing = GetListItemSpacing(listStyle, leading);
                                var listItems = new System.Collections.Generic.List<ColListItem>(nl2.Items.Count);
                                for (int itemIndex = 0; itemIndex < nl2.Items.Count; itemIndex++) {
                                    string text = nl2.Items[itemIndex];
                                    string marker = (nl2.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + ".";
                                    var lines = WrapSimpleText(text, availableWidth, baseFont, size);
                                    double firstLineWidth = lines.Count > 0 ? EstimateSimpleTextWidth(lines[0], baseFont, size) : 0;
                                    double firstLineDx = 0;
                                    if (nl2.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                                    else if (nl2.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);
                                    double spacingBefore = itemIndex == 0 ? listStyle?.SpacingBefore ?? 0D : 0D;
                                    double spacingAfter = itemIndex == nl2.Items.Count - 1 ? listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing : itemSpacing;
                                    listItems.Add(new ColListItem { Lines = lines, Marker = marker, MarkerXOffset = listLeftIndent + firstLineDx, MarkerWidth = markerWidth, MarkerAlign = PdfAlign.Right, TextXOffset = listLeftIndent + indent, TextWidth = alignmentWidth, TextAlign = nl2.Align, Color = nl2.Color ?? listStyle?.Color, Leading = leading, Size = size, SpacingBefore = spacingBefore, SpacingAfter = spacingAfter });
                                }

                                if ((listStyle?.KeepTogether == true || listStyle?.KeepWithNext == true) && listItems.Count > 0) {
                                    double listGroupHeight = 0D;
                                    foreach (var listItem in listItems) {
                                        listGroupHeight += listItem.SpacingBefore + listItem.Lines.Count * listItem.Leading + listItem.SpacingAfter;
                                    }

                                    if (listStyle?.KeepTogether == true) {
                                        listItems[0].IsFirstInKeepGroup = true;
                                        foreach (var listItem in listItems) {
                                            listItem.KeepTogether = true;
                                            listItem.KeepGroupHeight = listGroupHeight;
                                        }
                                    }

                                    if (listStyle?.KeepWithNext == true) {
                                        listItems[0].IsFirstInKeepWithNextGroup = true;
                                        foreach (var listItem in listItems) {
                                            listItem.KeepWithNext = true;
                                            listItem.KeepWithNextGroupItemCount = listItems.Count;
                                            listItem.KeepWithNextGroupHeight = listGroupHeight;
                                        }
                                    }
                                }

                                items.AddRange(listItems);
                            } else if (cb is PanelParagraphBlock ppb2) {
                                double size = currentOpts.DefaultFontSize;
                                double leading = size * 1.4;
                                var panelFont = ChooseNormal(currentOpts.DefaultFont);
                                double firstBaselineOffset = GetAscender(panelFont, size);
                                PanelStyle panelStyle = ResolvePanelStyle(ppb2, currentOpts);
                                double innerWidth = panelStyle.MaxWidth.HasValue ? Math.Min(colWs[i], panelStyle.MaxWidth.Value) : colWs[i];
                                ValidatePanelStyle(panelStyle, innerWidth);
                                double textWidthAvail = innerWidth - 2 * panelStyle.PaddingX;
                                var wrap = WrapRichRuns(ppb2.Runs, textWidthAvail, size, panelFont, leading);
                                double xOffset = 0;
                                if (panelStyle.Align == PdfAlign.Center) xOffset = Math.Max(0, (colWs[i] - innerWidth) / 2);
                                else if (panelStyle.Align == PdfAlign.Right) xOffset = Math.Max(0, colWs[i] - innerWidth);
                                items.Add(new ColPanel { Block = ppb2, Style = panelStyle, Lines = wrap.Lines, Heights = wrap.LineHeights, Leading = leading, Size = size, FirstBaselineOffset = firstBaselineOffset, XOffset = xOffset, PanelWidth = innerWidth, TextWidth = textWidthAvail });
                            } else if (cb is TableBlock tb2) {
                                var style = tb2.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
                                int cols = GetTableColumnCount(tb2);
                                if (cols == 0) {
                                    continue;
                                }

                                double padLeft = GetTableCellPaddingLeft(style);
                                double padRight = GetTableCellPaddingRight(style);
                                double padTop = GetTableCellPaddingTop(style);
                                double padBottom = GetTableCellPaddingBottom(style);
                                double size = currentOpts.DefaultFontSize;
                                ValidateTableRoleRowCounts(style, tb2.Rows.Count);
                                int headerRowCount = style.HeaderRowCount;
                                int footerRowCount = style.FooterRowCount;
                                int footerStartRowIndex = tb2.Rows.Count - footerRowCount;
                                ValidateTableCellStyleCoordinates(style, tb2.Rows.Count, cols);
                                ValidateTableColumnStyleBounds(style, cols);
                                ValidateTableRowSpansWithinRoleBoundaries(tb2, cols, headerRowCount, footerStartRowIndex);
                                double[]? autoFitWeights = style.AutoFitColumns
                                    ? MeasureAutoFitColumnWeights(tb2, currentOpts, style, size, headerRowCount, footerStartRowIndex)
                                    : null;
                                double[]? autoFitMinimumWidths = style.AutoFitColumns
                                    ? MeasureAutoFitColumnMinimumWidths(tb2, currentOpts, style, size, headerRowCount, footerStartRowIndex)
                                    : null;
                                double[] colPixel = new double[cols];
                                double[] colWeights = new double[cols];
                                bool[] fixedColumns = new bool[cols];
                                double?[] minWidths = new double?[cols];
                                double?[] maxWidths = new double?[cols];
                                double fixedWidthTotal = 0;
                                double totalWeight = 0;
                                for (int c = 0; c < cols; c++) {
                                    double? minWidth = GetOptionalColumnWidth(style.ColumnMinWidthPoints, c, "Table minimum column widths must be positive finite values.");
                                    if (!minWidth.HasValue && autoFitMinimumWidths != null && c < autoFitMinimumWidths.Length) {
                                        minWidth = autoFitMinimumWidths[c];
                                    }

                                    double? maxWidth = GetOptionalColumnWidth(style.ColumnMaxWidthPoints, c, "Table maximum column widths must be positive finite values.");
                                    if (minWidth.HasValue && maxWidth.HasValue && minWidth.Value > maxWidth.Value + 0.001) {
                                        throw new ArgumentException("Table minimum column widths cannot be greater than maximum column widths.");
                                    }

                                    minWidths[c] = minWidth;
                                    maxWidths[c] = maxWidth;

                                    if (style.ColumnWidthPoints != null && c < style.ColumnWidthPoints.Count && style.ColumnWidthPoints[c].HasValue) {
                                        double fixedWidth = style.ColumnWidthPoints[c]!.Value;
                                        if (minWidth.HasValue && fixedWidth < minWidth.Value - 0.001) {
                                            throw new ArgumentException("Table fixed column widths cannot be smaller than configured minimum widths.");
                                        }
                                        if (maxWidth.HasValue && fixedWidth > maxWidth.Value + 0.001) {
                                            throw new ArgumentException("Table fixed column widths cannot be greater than configured maximum widths.");
                                        }

                                        colPixel[c] = fixedWidth;
                                        fixedColumns[c] = true;
                                        fixedWidthTotal += fixedWidth;
                                        continue;
                                    }

                                    double weight = 1;
                                    if (style.ColumnWidthWeights != null && c < style.ColumnWidthWeights.Count) {
                                        weight = style.ColumnWidthWeights[c];
                                    } else if (autoFitWeights != null && c < autoFitWeights.Length) {
                                        weight = autoFitWeights[c];
                                    }

                                    colWeights[c] = weight;
                                    totalWeight += weight;
                                }

                                double tableAvailableWidth = ResolveTableAvailableWidth(style, colWs[i]);
                                if (fixedWidthTotal > tableAvailableWidth + 0.001) {
                                    throw new ArgumentException("Table fixed column widths exceed the available table width.");
                                }

                                double remainingWidth = Math.Max(0, tableAvailableWidth - fixedWidthTotal);
                                if (totalWeight <= 0) {
                                    remainingWidth = 0;
                                }

                                DistributeFlexibleColumns(colPixel, colWeights, fixedColumns, minWidths, maxWidths, remainingWidth);
                                double tableWidth = Math.Min(tableAvailableWidth, colPixel.Sum());
                                for (int c = 0; c < cols; c++) {
                                    if (colPixel[c] - padLeft - padRight <= 0.001) {
                                        throw new ArgumentException("Table horizontal cell padding must leave a positive text width.");
                                    }
                                }

                                var rowLines = new string[tb2.Rows.Count][][];
                                var rowLineCounts = new int[tb2.Rows.Count];
                                var rowHeights = new double[tb2.Rows.Count];
                                var rowLeadings = new double[tb2.Rows.Count];
                                var rowSizes = new double[tb2.Rows.Count];
                                var rowBold = new bool[tb2.Rows.Count];
                                for (int ri = 0; ri < tb2.Rows.Count; ri++) {
                                    double rowSize = GetTableRowFontSize(style, ri, headerRowCount, footerStartRowIndex, currentOpts.DefaultFontSize);
                                    double rowLeading = GetTableLeading(style, rowSize);
                                    bool rowUsesBold = GetTableRowBold(style, ri, headerRowCount, footerStartRowIndex);
                                    rowSizes[ri] = rowSize;
                                    rowLeadings[ri] = rowLeading;
                                    rowBold[ri] = rowUsesBold;
                                    rowLines[ri] = new string[cols][];
                                    int maxLines = 1;
                                    for (int ci = 0; ci < cols; ci++) {
                                        rowLines[ri][ci] = System.Array.Empty<string>();
                                    }

                                    var cells = GetTableCellLayouts(tb2, ri, cols);
                                    for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                                        TableCellLayout cell = cells[cellIndex];
                                        var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                                        double cellWidth = GetTableCellWidth(colPixel, cell.Column, cell.ColumnSpan, 0D);
                                        double innerWidth = Math.Max(1, cellWidth - padLeft - padRight);
                                        var lines = WrapSimpleText(cell.Text, innerWidth, cellFont, rowSize);
                                        rowLines[ri][cell.Column] = lines.ToArray();
                                        if (cell.RowSpan <= 1) {
                                            maxLines = Math.Max(maxLines, lines.Count);
                                        }
                                    }

                                    rowLineCounts[ri] = maxLines;
                                    rowHeights[ri] = Math.Max(maxLines * rowLeading + padTop + padBottom, style.MinRowHeight);
                                }
                                ApplyTableRowSpanHeights(tb2, cols, rowLines, rowHeights, rowLeadings, padTop, padBottom);

                                System.Collections.Generic.List<string>? captionLines = null;
                                double captionLeading = 0;
                                double captionHeight = 0;
                                if (!string.IsNullOrWhiteSpace(style.Caption)) {
                                    double captionSize = style.CaptionFontSize ?? size;
                                    captionLeading = captionSize * 1.25;
                                    var captionFont = ChooseNormal(currentOpts.DefaultFont);
                                    captionLines = WrapSimpleText(style.Caption!, tableWidth, captionFont, captionSize);
                                    captionHeight = captionLines.Count * captionLeading + style.CaptionSpacingAfter;
                                    double firstRowHeight = rowHeights.Length > 0 ? rowHeights[0] : 0;
                                    double maxContentHeightForCaption = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                    if (captionHeight + firstRowHeight > maxContentHeightForCaption + 0.001) {
                                        throw new ArgumentException("Table caption and first row exceed the available page content height.");
                                    }
                                }

                                items.Add(new ColTable { Block = tb2, Style = style, Columns = cols, ColumnWidths = colPixel, RowLines = rowLines, RowLineCounts = rowLineCounts, RowHeights = rowHeights, RowLeadings = rowLeadings, RowSizes = rowSizes, RowBold = rowBold, Width = tableWidth, Size = size, HeaderRowCount = headerRowCount, FooterStartRowIndex = footerStartRowIndex, CaptionLines = captionLines, CaptionLeading = captionLeading, CaptionHeight = captionHeight });
                            } else if (cb is HorizontalRuleBlock hr2) {
                                items.Add(new ColRule { Block = hr2 });
                            } else if (cb is ImageBlock ib2) {
                                items.Add(new ColImg { Block = ib2 });
                            } else if (cb is ShapeBlock sb2) {
                                items.Add(new ColShape { Block = sb2 });
                            } else if (cb is DrawingBlock db2) {
                                items.Add(new ColDrawing { Block = db2 });
                            } else if (cb is BookmarkBlock bookmark2) {
                                items.Add(new ColBookmark { Block = bookmark2 });
                            } else if (cb is SpacerBlock spacer2) {
                                items.Add(new ColSpacer { Block = spacer2 });
                            }
                        }
                        colItems.Add(items);
                    }

                    double MeasureRowKeepTogetherHeight(System.Collections.Generic.List<ColItem> items) {
                        double total = 0D;
                        foreach (var item in items) {
                            if (item is ColPar paragraph) {
                                PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph.Block);
                                total += GetParagraphSpacingBefore(paragraphStyle) + paragraph.Heights.Sum() + GetParagraphSpacingAfter(paragraphStyle, paragraph.Leading);
                            } else if (item is ColHead heading) {
                                total += heading.SpacingBefore + heading.Lines.Count * heading.Leading + heading.SpacingAfter;
                            } else if (item is ColListItem listItem) {
                                total += listItem.SpacingBefore + listItem.Lines.Count * listItem.Leading + listItem.SpacingAfter;
                            } else if (item is ColPanel panel) {
                                total += panel.Style.SpacingBefore + panel.Style.PaddingY + panel.Heights.Sum() + panel.Style.PaddingY + panel.Style.SpacingAfter;
                            } else if (item is ColTable table) {
                                total += table.Style.SpacingBefore + table.CaptionHeight + table.RowHeights.Sum() + table.Style.SpacingAfter;
                            } else if (item is ColRule rule) {
                                PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(rule.Block, currentOpts);
                                ValidateHorizontalRule(ruleStyle);
                                total += ruleStyle.SpacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
                            } else if (item is ColImg image) {
                                PdfImageStyle imageStyle = ResolveImageStyle(image.Block, currentOpts);
                                total += imageStyle.SpacingBefore + image.Block.Height + imageStyle.SpacingAfter;
                            } else if (item is ColShape shape) {
                                PdfDrawingStyle shapeStyle = ResolveDrawingStyle(shape.Block, currentOpts);
                                total += shapeStyle.SpacingBefore + shape.Block.Shape.Height + shapeStyle.SpacingAfter;
                            } else if (item is ColDrawing drawing) {
                                PdfDrawingStyle drawingStyle = ResolveDrawingStyle(drawing.Block, currentOpts);
                                total += drawingStyle.SpacingBefore + drawing.Block.Drawing.Height + drawingStyle.SpacingAfter;
                            } else if (item is ColSpacer spacerItem) {
                                total += spacerItem.Block.Height;
                            }
                        }

                        return total;
                    }

                    double MeasureColItemFirstVisualHeight(ColItem item) {
                        if (item is ColPar paragraph) {
                            PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(paragraph.Block);
                            return GetParagraphSpacingBefore(paragraphStyle) + (paragraph.Heights.Count == 0 ? 0D : paragraph.Heights[0]);
                        }

                        if (item is ColHead heading) {
                            return heading.SpacingBefore + (heading.Lines.Count == 0 ? 0D : heading.Leading);
                        }

                        if (item is ColListItem listItem) {
                            return listItem.SpacingBefore + (listItem.Lines.Count == 0 ? 0D : listItem.Leading);
                        }

                        if (item is ColPanel panel) {
                            return panel.Style.SpacingBefore + panel.Style.PaddingY + (panel.Heights.Count == 0 ? 0D : panel.Heights[0]) + panel.Style.PaddingY;
                        }

                        if (item is ColTable table) {
                            double firstRowHeight = table.RowHeights.Length == 0 ? 0D : table.RowHeights[0];
                            return table.Style.SpacingBefore + table.CaptionHeight + firstRowHeight;
                        }

                        if (item is ColRule rule) {
                            PdfHorizontalRuleStyle ruleStyle = ResolveHorizontalRuleStyle(rule.Block, currentOpts);
                            return ruleStyle.SpacingBefore + ruleStyle.Thickness + ruleStyle.SpacingAfter;
                        }

                        if (item is ColImg image) {
                            PdfImageStyle imageStyle = ResolveImageStyle(image.Block, currentOpts);
                            return imageStyle.SpacingBefore + image.Block.Height + imageStyle.SpacingAfter;
                        }

                        if (item is ColShape shape) {
                            PdfDrawingStyle shapeStyle = ResolveDrawingStyle(shape.Block, currentOpts);
                            return shapeStyle.SpacingBefore + shape.Block.Shape.Height + shapeStyle.SpacingAfter;
                        }

                        if (item is ColDrawing drawing) {
                            PdfDrawingStyle drawingStyle = ResolveDrawingStyle(drawing.Block, currentOpts);
                            return drawingStyle.SpacingBefore + drawing.Block.Drawing.Height + drawingStyle.SpacingAfter;
                        }

                        if (item is ColSpacer spacerItem) {
                            return spacerItem.Block.Height;
                        }

                        return 0D;
                    }

                    double? rowContentHeightCache = null;
                    double GetRowContentHeight() {
                        if (rowContentHeightCache.HasValue) {
                            return rowContentHeightCache.Value;
                        }

                        double measuredHeight = 0D;
                        foreach (var items in colItems) {
                            measuredHeight = Math.Max(measuredHeight, MeasureRowKeepTogetherHeight(items));
                        }

                        rowContentHeightCache = measuredHeight;
                        return measuredHeight;
                    }

                    if (rowStyle?.KeepTogether == true) {
                        double rowContentHeight = GetRowContentHeight();
                        double rowKeepHeight = rowSpacingBefore + rowContentHeight + rowSpacingAfter;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (rowKeepHeight > availableHeight + 0.001) {
                            throw new ArgumentException("Row height exceeds the available page content height.");
                        }

                        if (y < yStart - 0.001 && y - rowKeepHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    if (rowStyle?.KeepWithNext == true && nextBlock != null) {
                        double rowContentHeight = GetRowContentHeight();
                        double rowHeight = rowSpacingBefore + rowContentHeight + rowSpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        double keepHeight = rowHeight + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && rowHeight <= availableHeight + 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    if (rowSpacingBefore > 0) {
                        if (y - rowSpacingBefore < currentOpts.MarginBottom) {
                            NewPage();
                        }

                        y -= rowSpacingBefore;
                    }

                    bool AnyRemaining() {
                        for (int i = 0; i < ncols; i++) if (colStates[i].idx < colItems[i].Count) return true; return false;
                    }

                    int rowColumnFlowGuard = 0;
                    while (AnyRemaining()) {
                        rowColumnFlowGuard++;
                        if (rowColumnFlowGuard > 10000) {
                            throw new InvalidOperationException("Row column layout did not make forward progress.");
                        }

                        double avail = y - currentOpts.MarginBottom;
                        if (avail <= 0.5) { NewPage(); avail = y - currentOpts.MarginBottom; }

                        double maxConsumed = 0;
                        bool anyColumnAdvanced = false;
                        for (int ci = 0; ci < ncols; ci++) {
                            var items = colItems[ci];
                            var (idx, line, subline) = colStates[ci];
                            var startState = (idx, line, subline);
                            double xCol = colXs[ci];
                            double wCol = colWs[ci];
                            double yCol = y;
                            double consumed = 0;
                            double remain = avail;
                            while (idx < items.Count && remain > 0.1) {
                                var it = items[idx];
                                if (it is ColPar par) {
                                    var pblock = par.Block;
                                    var lines = par.Lines;
                                    var heights = par.Heights;
                                    double leading = par.Leading;
                                    double size = par.Size;
                                    PdfParagraphStyle? paragraphStyle = EffectiveParagraphStyle(pblock);
                                    double spacingBefore = line == 0 && consumed > 0.001 ? GetParagraphSpacingBefore(paragraphStyle) : 0;
                                    double spacingAfter = GetParagraphSpacingAfter(paragraphStyle, leading);
                                    if (paragraphStyle?.KeepWithNext == true && line == 0 && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = spacingBefore + heights.Sum() + spacingAfter + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (paragraphStyle?.KeepTogether == true && line == 0) {
                                        double paragraphHeight = spacingBefore + heights.Sum() + spacingAfter;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (paragraphHeight > availableHeight + 0.001) {
                                            throw new ArgumentException("Paragraph height exceeds the available page content height.");
                                        }

                                        if (paragraphHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    double availableForLines = remain - spacingBefore;
                                    if (availableForLines < 0) {
                                        if (consumed > 0) break;
                                        remain = 0;
                                        break;
                                    }

                                    int start = line;
                                    int take = 0; double hsum = 0;
                                    for (int li2 = start; li2 < lines.Count; li2++) {
                                        double hAdd = heights[li2];
                                        if (hsum + hAdd + (li2 == lines.Count - 1 ? spacingAfter : 0) > availableForLines) break;
                                        hsum += hAdd; take++;
                                    }

                                    if (TryApplyWidowControl(paragraphStyle, lines.Count, start, ref take, ref hsum, heights, consumed > 0 || y < yStart - 0.001)) {
                                        break;
                                    }

                                    if (take == 0) break;
                                    if (spacingBefore > 0) {
                                        yCol -= spacingBefore;
                                        remain -= spacingBefore;
                                        consumed += spacingBefore;
                                    }

                                    var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                                    var sliceHeights = new System.Collections.Generic.List<double>();
                                    for (int k = 0; k < take; k++) { sliceLines.Add(lines[start + k]); sliceHeights.Add(heights[start + k]); }
                                    pageDirty = true;
                                    var paragraphFont = ChooseNormal(currentOpts.DefaultFont);
                                    WriteRichParagraph(sb, pblock, sliceLines, sliceHeights, currentOpts, FirstTextBaselineFromTop(paragraphFont, size, yCol), size, leading, currentPage!.Annotations, xCol + par.XOffset, par.TextWidth, start == 0 ? xCol + par.FirstLineXOffset : null, start == 0 ? par.FirstLineTextWidth : null);
                                    if (pblock.Runs.Any(r => r.Bold)) { currentPage!.UsedBold = true; usedBold = true; }
                                    if (pblock.Runs.Any(r => r.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
                                    if (pblock.Runs.Any(r => r.Bold && r.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }
                                    yCol -= hsum; remain -= hsum; consumed += hsum; line += take;
                                    if (line >= lines.Count) { double space = spacingAfter; if (space <= remain) { yCol -= space; remain -= space; consumed += space; } idx++; line = 0; }
                                } else if (it is ColHead ch) {
                                    var hb2 = ch.Block;
                                    var lines = ch.Lines;
                                    double leading = ch.Leading;
                                    double size = ch.Size;
                                    double needed = ch.SpacingBefore + lines.Count * leading + ch.SpacingAfter;
                                    if (ch.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = needed + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    if (ch.SpacingBefore > 0) {
                                        yCol -= ch.SpacingBefore;
                                        remain -= ch.SpacingBefore;
                                        consumed += ch.SpacingBefore;
                                    }

                                    if (currentOpts.CreateOutlineFromHeadings) {
                                        currentPage!.Bookmarks.Add(new PageBookmark { Level = hb2.Level, Title = hb2.Text, Y = yCol });
                                    }
                                    var headingFont = ChooseBold(ChooseNormal(currentOpts.DefaultFont));
                                    double firstBaseline = FirstTextBaselineFromTop(headingFont, size, yCol);
                                    AddHeadingLinkAnnotations(hb2, lines, headingFont, size, leading, xCol, wCol, firstBaseline);
                                    WriteLinesInternal("F2", size, leading, xCol, wCol, firstBaseline, lines, hb2.Align, ch.Color, applyBaselineTweak: false);
                                    currentPage!.UsedBold = true;
                                    usedBold = true;
                                    double consumedHeight = lines.Count * leading + ch.SpacingAfter;
                                    yCol -= consumedHeight; remain -= consumedHeight; consumed += consumedHeight; idx++;
                                } else if (it is ColListItem listItem) {
                                    var lines = listItem.Lines;
                                    double leading = listItem.Leading;
                                    if (line == 0 && listItem.KeepTogether && listItem.IsFirstInKeepGroup) {
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (listItem.KeepGroupHeight > availableHeight + 0.001) {
                                            throw new ArgumentException("List height exceeds the available page content height.");
                                        }

                                        if (listItem.KeepGroupHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (line == 0 && listItem.KeepWithNext && listItem.IsFirstInKeepWithNextGroup) {
                                        int nextItemIndex = idx + listItem.KeepWithNextGroupItemCount;
                                        if (nextItemIndex < items.Count) {
                                            double nextHeight = MeasureColItemFirstVisualHeight(items[nextItemIndex]);
                                            double keepHeight = listItem.KeepWithNextGroupHeight + nextHeight;
                                            double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                            if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                                if (consumed > 0) break;
                                                remain = 0;
                                                break;
                                            }
                                        }
                                    }

                                    if (line == 0 && listItem.SpacingBefore > 0) {
                                        if (listItem.SpacingBefore > remain && consumed > 0) break;
                                        if (listItem.SpacingBefore > remain && consumed == 0) { remain = 0; break; }
                                        yCol -= listItem.SpacingBefore;
                                        remain -= listItem.SpacingBefore;
                                        consumed += listItem.SpacingBefore;
                                    }

                                    double availableForLines = remain;
                                    int start = line;
                                    int take = 0;
                                    double hsum = 0;
                                    for (int li2 = start; li2 < lines.Count; li2++) {
                                        if (hsum + leading > availableForLines) break;
                                        hsum += leading;
                                        take++;
                                    }
                                    if (take == 0) break;

                                    var sliceLines = new System.Collections.Generic.List<string>(take);
                                    for (int k = 0; k < take; k++) {
                                        sliceLines.Add(lines[start + k]);
                                    }

                                    pageDirty = true;
                                    var listFont = ChooseNormal(currentOpts.DefaultFont);
                                    double baselineY = FirstTextBaselineFromTop(listFont, listItem.Size, yCol);
                                    if (line == 0) {
                                        var markerLines = new System.Collections.Generic.List<string>(1) { listItem.Marker };
                                        WriteLinesInternal("F1", listItem.Size, leading, xCol + listItem.MarkerXOffset, listItem.MarkerWidth, baselineY, markerLines, listItem.MarkerAlign, listItem.Color, applyBaselineTweak: true);
                                    }

                                    WriteLinesInternal("F1", listItem.Size, leading, xCol + listItem.TextXOffset, listItem.TextWidth, baselineY, sliceLines, listItem.TextAlign, listItem.Color, applyBaselineTweak: true);
                                    yCol -= hsum;
                                    remain -= hsum;
                                    consumed += hsum;
                                    line += take;
                                    if (line >= lines.Count) {
                                        double space = listItem.SpacingAfter;
                                        if (space <= remain) {
                                            yCol -= space;
                                            remain -= space;
                                            consumed += space;
                                        }

                                        idx++;
                                        line = 0;
                                    }
                                } else if (it is ColPanel panel) {
                                    var pblock = panel.Block;
                                    var panelStyle = panel.Style;
                                    var lines = panel.Lines;
                                    var heights = panel.Heights;
                                    double xPanel = xCol + panel.XOffset;
                                    if (line == 0 && panelStyle.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double panelHeight = panelStyle.SpacingBefore + panelStyle.PaddingY + heights.Sum() + panelStyle.PaddingY + panelStyle.SpacingAfter;
                                        double keepHeight = panelHeight + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (line == 0 && panelStyle.SpacingBefore > 0) {
                                        if (panelStyle.SpacingBefore > remain && consumed > 0) break;
                                        if (panelStyle.SpacingBefore > remain && consumed == 0) { remain = 0; break; }
                                        yCol -= panelStyle.SpacingBefore;
                                        remain -= panelStyle.SpacingBefore;
                                        consumed += panelStyle.SpacingBefore;
                                    }

                                    if (panelStyle.KeepTogether) {
                                        double textHeight = heights.Sum();
                                        double panelHeight = panelStyle.PaddingY + textHeight + panelStyle.PaddingY;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (panelHeight > availableHeight + 0.001) {
                                            throw new ArgumentException("Panel height exceeds the available page content height.");
                                        }

                                        if (panelHeight > remain && consumed > 0) break;
                                        if (panelHeight > remain && consumed == 0) { remain = 0; break; }

                                        double panelTop = yCol;
                                        double panelBottom = yCol - panelHeight;
                                        if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom); }
                                        if (panelStyle.BorderColor.HasValue && panelStyle.BorderWidth > 0) { pageDirty = true; DrawRowRect(sb, panelStyle.BorderColor.Value, panelStyle.BorderWidth, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom); }
                                        pageDirty = true;
                                        WriteRichParagraph(sb, new RichParagraphBlock(pblock.Runs, pblock.Align, pblock.DefaultColor), lines, heights, currentOpts, panelTop - panelStyle.PaddingY - panel.FirstBaselineOffset, panel.Size, panel.Leading, currentPage!.Annotations, xPanel + panelStyle.PaddingX, panel.TextWidth);
                                        if (pblock.Runs.Any(r => r.Bold)) { currentPage!.UsedBold = true; usedBold = true; }
                                        if (pblock.Runs.Any(r => r.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
                                        if (pblock.Runs.Any(r => r.Bold && r.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }

                                        yCol = panelBottom;
                                        remain -= panelHeight;
                                        consumed += panelHeight;
                                        if (panelStyle.SpacingAfter > 0 && panelStyle.SpacingAfter <= remain) {
                                            yCol -= panelStyle.SpacingAfter;
                                            remain -= panelStyle.SpacingAfter;
                                            consumed += panelStyle.SpacingAfter;
                                        }
                                        idx++;
                                        line = 0;
                                    } else {
                                        int start = line;
                                        double topPad = start == 0 ? panelStyle.PaddingY : 0;
                                        double minLine = heights[start];
                                        if (remain < topPad + minLine) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }

                                        double roomForText = remain - topPad - panelStyle.PaddingY;
                                        if (roomForText < minLine) {
                                            roomForText = remain - topPad;
                                        }

                                        int take = 0;
                                        double hsum = 0;
                                        for (int k = start; k < lines.Count; k++) {
                                            double h = heights[k];
                                            if (hsum + h > roomForText) break;
                                            hsum += h;
                                            take++;
                                        }

                                        if (take == 0) break;

                                        bool lastSeg = start + take >= lines.Count;
                                        double panelTop = yCol;
                                        double usedBottomPad = lastSeg ? panelStyle.PaddingY : Math.Max(0, remain - (topPad + hsum));
                                        double panelBottom = yCol - (topPad + hsum + usedBottomPad);
                                        if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom); }
                                        if (panelStyle.BorderColor.HasValue && panelStyle.BorderWidth > 0) { pageDirty = true; DrawRowRect(sb, panelStyle.BorderColor.Value, panelStyle.BorderWidth, xPanel, panelBottom, panel.PanelWidth, panelTop - panelBottom); }

                                        var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                                        var sliceHeights = new System.Collections.Generic.List<double>();
                                        for (int k = 0; k < take; k++) {
                                            sliceLines.Add(lines[start + k]);
                                            sliceHeights.Add(heights[start + k]);
                                        }

                                        pageDirty = true;
                                        WriteRichParagraph(sb, new RichParagraphBlock(pblock.Runs, pblock.Align, pblock.DefaultColor), sliceLines, sliceHeights, currentOpts, panelTop - topPad - panel.FirstBaselineOffset, panel.Size, panel.Leading, currentPage!.Annotations, xPanel + panelStyle.PaddingX, panel.TextWidth);
                                        if (pblock.Runs.Any(r => r.Bold)) { currentPage!.UsedBold = true; usedBold = true; }
                                        if (pblock.Runs.Any(r => r.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
                                        if (pblock.Runs.Any(r => r.Bold && r.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }

                                        double segmentHeight = panelTop - panelBottom;
                                        yCol = panelBottom;
                                        remain -= segmentHeight;
                                        consumed += segmentHeight;
                                        line += take;
                                        if (line >= lines.Count) {
                                            if (panelStyle.SpacingAfter > 0 && panelStyle.SpacingAfter <= remain) {
                                                yCol -= panelStyle.SpacingAfter;
                                                remain -= panelStyle.SpacingAfter;
                                                consumed += panelStyle.SpacingAfter;
                                            }
                                            idx++;
                                            line = 0;
                                        } else {
                                            break;
                                        }
                                    }
                                } else if (it is ColTable table) {
                                    var tbColumn = table.Block;
                                    var tableStyle = table.Style;
                                    double padLeft = GetTableCellPaddingLeft(tableStyle);
                                    double padRight = GetTableCellPaddingRight(tableStyle);
                                    double padTop = GetTableCellPaddingTop(tableStyle);
                                    double padBottom = GetTableCellPaddingBottom(tableStyle);
                                    double xTable = ResolveTableX(tbColumn.Align, tableStyle, xCol, wCol, table.Width);

                                    double maxContentHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                    if (line == 0 && tableStyle.KeepTogether) {
                                        double keepHeight = tableStyle.SpacingBefore + table.CaptionHeight + table.RowHeights.Sum() + tableStyle.SpacingAfter;
                                        if (keepHeight > maxContentHeight + 0.001) {
                                            throw new ArgumentException("Table height exceeds the available page content height.");
                                        }

                                        if (keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (line == 0 && tableStyle.KeepWithNext && idx + 1 < items.Count) {
                                        double tableHeight = tableStyle.SpacingBefore + table.CaptionHeight + table.RowHeights.Sum() + tableStyle.SpacingAfter;
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = tableHeight + nextHeight;
                                        if (nextHeight > 0.001 && tableHeight <= maxContentHeight + 0.001 && keepHeight <= maxContentHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (line == 0 && tableStyle.SpacingBefore > 0) {
                                        if (tableStyle.SpacingBefore > remain && consumed > 0) break;
                                        if (tableStyle.SpacingBefore > remain && consumed == 0) { remain = 0; break; }
                                        yCol -= tableStyle.SpacingBefore;
                                        remain -= tableStyle.SpacingBefore;
                                        consumed += tableStyle.SpacingBefore;
                                    }

                                    if (line == 0 && table.CaptionLines != null) {
                                        double firstRowHeight = table.RowHeights.Length > 0 ? table.RowHeights[0] : 0;
                                        double neededWithFirstRow = table.CaptionHeight + firstRowHeight;
                                        if (neededWithFirstRow > maxContentHeight + 0.001) {
                                            throw new ArgumentException("Table caption and first row exceed the available page content height.");
                                        }
                                        if (neededWithFirstRow > remain && consumed > 0) break;
                                        if (neededWithFirstRow > remain && consumed == 0) { remain = 0; break; }

                                        double captionSize = tableStyle.CaptionFontSize ?? table.Size;
                                        var captionFont = ChooseNormal(currentOpts.DefaultFont);
                                        pageDirty = true;
                                        WriteLinesInternal("F1", captionSize, table.CaptionLeading, xTable, table.Width, yCol - GetAscender(captionFont, captionSize), table.CaptionLines, tableStyle.CaptionAlign, tableStyle.CaptionColor);
                                        yCol -= table.CaptionHeight;
                                        remain -= table.CaptionHeight;
                                        consumed += table.CaptionHeight;
                                    }

                                    double repeatHeaderHeight = 0;
                                    for (int headerIndex = 0; headerIndex < table.HeaderRowCount; headerIndex++) {
                                        repeatHeaderHeight += table.RowHeights[headerIndex];
                                    }

                                    bool HasRepeatableHeader() =>
                                        table.HeaderRowCount > 0 &&
                                        tbColumn.Rows.Count > table.HeaderRowCount;

                                    bool AtContinuationPageTop() =>
                                        Math.Abs(yCol - yStart) <= 0.001;

                                    void DrawColumnTableRowSegment(int rowIndex, bool renderAsHeader, int startLine, int lineCount) {
                                        bool renderAsFooter = rowIndex >= table.FooterStartRowIndex;
                                        bool rowUsesBold = table.RowBold[rowIndex];
                                        double rowSize = table.RowSizes[rowIndex];
                                        double rowLeading = table.RowLeadings[rowIndex];
                                        bool wholeRowSegment = startLine == 0 && lineCount == table.RowLineCounts[rowIndex];
                                        double rowHeight = wholeRowSegment ? table.RowHeights[rowIndex] : Math.Max(1, lineCount) * rowLeading + padTop + padBottom;
                                        if (rowUsesBold) {
                                            currentPage!.UsedBold = true;
                                            usedBold = true;
                                        }

                                        var cells = GetTableCellLayouts(tbColumn, rowIndex, table.Columns);
                                        double rowBottom = yCol - rowHeight;
                                        int bodyRowIndex = rowIndex - table.HeaderRowCount;
                                        bool stripeBodyRow = bodyRowIndex >= 0 && bodyRowIndex % 2 == 1;
                                        bool[] rowFillSkips = GetRowSpanContinuationSkipColumns(tbColumn, rowIndex, table.Columns);
                                        if (tableStyle.HeaderFill is not null && renderAsHeader) { pageDirty = true; DrawTableRowFill(sb, tableStyle.HeaderFill.Value, xTable, table.ColumnWidths, 0D, rowBottom, rowHeight, rowFillSkips); }
                                        else if (tableStyle.FooterFill is not null && renderAsFooter) { pageDirty = true; DrawTableRowFill(sb, tableStyle.FooterFill.Value, xTable, table.ColumnWidths, 0D, rowBottom, rowHeight, rowFillSkips); }
                                        else if (!renderAsHeader && !renderAsFooter && tableStyle.RowStripeFill is not null && stripeBodyRow) { pageDirty = true; DrawTableRowFill(sb, tableStyle.RowStripeFill.Value, xTable, table.ColumnWidths, 0D, rowBottom, rowHeight, rowFillSkips); }

                                        if (!renderAsHeader && !renderAsFooter && tableStyle.BodyColumnFills != null) {
                                            bool[] bodyColumnFillSkips = GetMergedCellContinuationSkipColumns(tbColumn, rowIndex, table.Columns);
                                            double fillX = xTable;
                                            for (int fillColumn = 0; fillColumn < table.Columns; fillColumn++) {
                                                PdfColor? fill = fillColumn < tableStyle.BodyColumnFills.Count ? tableStyle.BodyColumnFills[fillColumn] : null;
                                                if (fill.HasValue && (fillColumn >= bodyColumnFillSkips.Length || !bodyColumnFillSkips[fillColumn])) {
                                                    pageDirty = true;
                                                    DrawRowFill(sb, fill.Value, fillX, rowBottom, table.ColumnWidths[fillColumn], rowHeight);
                                                }
                                                fillX += table.ColumnWidths[fillColumn];
                                            }
                                        }

                                        if (tableStyle.CellFills != null && tableStyle.CellFills.Count > 0) {
                                            double fillX = xTable;
                                            for (int fillColumn = 0; fillColumn < table.Columns; fillColumn++) {
                                                if (tableStyle.CellFills.TryGetValue((rowIndex, fillColumn), out PdfColor fill) &&
                                                    TryGetTableCellLayoutAtColumn(cells, fillColumn, out TableCellLayout fillCell) &&
                                                    (fillColumn >= rowFillSkips.Length || !rowFillSkips[fillColumn])) {
                                                    int span = wholeRowSegment ? fillCell.ColumnSpan : 1;
                                                    double fillHeight = rowHeight;
                                                    double fillBottom = rowBottom;
                                                    if (wholeRowSegment) {
                                                        if (fillCell.RowSpan > 1) {
                                                            fillHeight = GetTableCellHeight(table.RowHeights, rowIndex, fillCell.RowSpan);
                                                            fillBottom = yCol - fillHeight;
                                                        }
                                                    }

                                                    pageDirty = true;
                                                    DrawRowFill(sb, fill, fillX, fillBottom, GetTableCellWidth(table.ColumnWidths, fillColumn, span, 0D), fillHeight);
                                                }
                                                fillX += table.ColumnWidths[fillColumn];
                                            }
                                        }

                                        string fontRes = GetTableRowFontResource(rowUsesBold);
                                        var textColor = renderAsHeader ? tableStyle.HeaderTextColor : renderAsFooter ? tableStyle.FooterTextColor : tableStyle.TextColor;
                                        double xi = xTable;
                                        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                                            TableCellLayout cell = cells[cellIndex];
                                            int c = cell.Column;
                                            xi = xTable;
                                            for (int xColumn = 0; xColumn < c; xColumn++) {
                                                xi += table.ColumnWidths[xColumn];
                                            }

                                            double cellWidth = GetTableCellWidth(table.ColumnWidths, c, cell.ColumnSpan, 0D);
                                            double innerW = cellWidth - padLeft - padRight;
                                            double cellHeight = wholeRowSegment && cell.RowSpan > 1 ? GetTableCellHeight(table.RowHeights, rowIndex, cell.RowSpan) : rowHeight;
                                            double cellBottom = yCol - cellHeight;
                                            var align = PdfColumnAlign.Left;
                                            if (tableStyle.Alignments != null && c < tableStyle.Alignments.Count) align = tableStyle.Alignments[c];
                                            if (tableStyle.RightAlignNumeric && LooksNumeric(cell.Text)) align = PdfColumnAlign.Right;
                                            var verticalAlign = PdfCellVerticalAlign.Top;
                                            if (tableStyle.VerticalAlignments != null && c < tableStyle.VerticalAlignments.Count) verticalAlign = tableStyle.VerticalAlignments[c];
                                            var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                                            var lines = table.RowLines[rowIndex][c];
                                            int sourceStartLine = wholeRowSegment && cell.RowSpan > 1 ? 0 : startLine;
                                            int requestedLineCount = wholeRowSegment && cell.RowSpan > 1 ? lines.Length : lineCount;
                                            int visibleLineCount = Math.Max(0, Math.Min(requestedLineCount, lines.Length - sourceStartLine));
                                            double verticalOffset = 0;
                                            if (visibleLineCount > 0) {
                                                double availableTextHeight = Math.Max(0, cellHeight - padTop - padBottom);
                                                double visibleTextHeight = visibleLineCount * rowLeading;
                                                double unusedTextHeight = Math.Max(0, availableTextHeight - visibleTextHeight);
                                                if (verticalAlign == PdfCellVerticalAlign.Middle) verticalOffset = unusedTextHeight / 2;
                                                else if (verticalAlign == PdfCellVerticalAlign.Bottom) verticalOffset = unusedTextHeight;
                                            }

                                            double firstBaseline = yCol - padTop - verticalOffset - GetAscender(cellFont, rowSize) + tableStyle.RowBaselineOffset;

                                            pageDirty = true;
                                            for (int lineIndex = 0; lineIndex < visibleLineCount; lineIndex++) {
                                                string cellLine = lines[sourceStartLine + lineIndex];
                                                double textW = EstimateSimpleTextWidth(cellLine, cellFont, rowSize);
                                                double offset = 0;
                                                if (align == PdfColumnAlign.Center) offset = Math.Max(0, (innerW - textW) / 2);
                                                else if (align == PdfColumnAlign.Right) offset = Math.Max(0, innerW - textW);
                                                double xCell = xi + padLeft + offset;
                                                double yCell = firstBaseline - lineIndex * rowLeading;
                                                if (ShouldClipTableCellText(xCell, yCell, textW, cellFont, rowSize, xi, cellBottom, cellWidth, cellHeight)) {
                                                    WriteClippedCell(sb, fontRes, rowSize, xCell, yCell, cellLine, textColor, currentOpts, xi - TableCellClipBleed, cellBottom - TableCellClipBleed, cellWidth + TableCellClipBleed * 2D, cellHeight + TableCellClipBleed * 2D);
                                                } else {
                                                    WriteCell(sb, fontRes, rowSize, xCell, yCell, cellLine, textColor, currentOpts);
                                                }
                                            }

                                            string? linkUri = cell.LinkUri;
                                            string? linkContents = cell.LinkContents;
                                            if (tbColumn.Links.TryGetValue((rowIndex, c), out var uri)) {
                                                linkUri = uri;
                                                linkContents = cell.Text;
                                            }

                                            if (!string.IsNullOrEmpty(linkUri)) {
                                                currentPage!.Annotations.Add(new LinkAnnotation { X1 = xi + padLeft, Y1 = cellBottom + padBottom, X2 = xi + padLeft + innerW, Y2 = yCol - padTop, Uri = linkUri!, Contents = linkContents ?? cell.Text });
                                            }
                                        }

                                        if (tableStyle.BorderColor is not null && tableStyle.BorderWidth > 0) {
                                            pageDirty = true;
                                            bool[] topBorderSkips = GetRowSpanBoundarySkipColumns(tbColumn, rowIndex - 1, table.Columns);
                                            bool[] bottomBorderSkips = GetRowSpanBoundarySkipColumns(tbColumn, rowIndex, table.Columns);
                                            bool segmentBorderRows = HasSkippedColumns(topBorderSkips, table.Columns) || HasSkippedColumns(bottomBorderSkips, table.Columns);
                                            if (segmentBorderRows) {
                                                DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, 0D, rowBottom + rowHeight, topBorderSkips);
                                                DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, 0D, rowBottom, bottomBorderSkips);
                                                DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, rowBottom + rowHeight, rowBottom);
                                                DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable + table.Width, rowBottom + rowHeight, rowBottom);
                                            } else {
                                                DrawRowRect(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, rowBottom, table.Width, rowHeight);
                                            }

                                            double xi2 = xTable;
                                            for (int c = 0; c < table.Columns - 1; c++) {
                                                xi2 += table.ColumnWidths[c];
                                                if (IsTableBoundaryInsideSpannedCell(tbColumn, rowIndex, c, table.Columns)) {
                                                    continue;
                                                }

                                                DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xi2, rowBottom + rowHeight, rowBottom);
                                            }
                                        }

                                        if (renderAsFooter && rowIndex == table.FooterStartRowIndex) {
                                            PdfColor? footerSeparatorColor = tableStyle.FooterSeparatorColor ?? tableStyle.RowSeparatorColor;
                                            double footerSeparatorWidth = tableStyle.FooterSeparatorWidth > 0 ? tableStyle.FooterSeparatorWidth : tableStyle.RowSeparatorWidth;
                                            if (footerSeparatorColor is not null && footerSeparatorWidth > 0) {
                                                pageDirty = true;
                                                DrawTableHorizontalLine(sb, footerSeparatorColor.Value, footerSeparatorWidth, xTable, table.ColumnWidths, 0D, yCol, GetRowSpanBoundarySkipColumns(tbColumn, rowIndex - 1, table.Columns));
                                            }
                                        }

                                        PdfColor? separatorColor = renderAsHeader && tableStyle.HeaderSeparatorColor is not null ? tableStyle.HeaderSeparatorColor : tableStyle.RowSeparatorColor;
                                        double separatorWidth = renderAsHeader && tableStyle.HeaderSeparatorWidth > 0 ? tableStyle.HeaderSeparatorWidth : tableStyle.RowSeparatorWidth;
                                        if (separatorColor is not null && separatorWidth > 0) {
                                            pageDirty = true;
                                            DrawTableHorizontalLine(sb, separatorColor.Value, separatorWidth, xTable, table.ColumnWidths, 0D, rowBottom, GetRowSpanBoundarySkipColumns(tbColumn, rowIndex, table.Columns));
                                        }

                                        if (tableStyle.CellBorders != null && tableStyle.CellBorders.Count > 0) {
                                            double borderX = xTable;
                                            for (int borderColumn = 0; borderColumn < table.Columns; borderColumn++) {
                                                if (tableStyle.CellBorders.TryGetValue((rowIndex, borderColumn), out PdfCellBorder? cellBorder) &&
                                                    TryGetTableCellLayoutAtColumn(cells, borderColumn, out TableCellLayout borderCell) &&
                                                    (borderColumn >= rowFillSkips.Length || !rowFillSkips[borderColumn]) &&
                                                    cellBorder != null &&
                                                    cellBorder.Color.HasValue &&
                                                    cellBorder.Width > 0) {
                                                    int span = wholeRowSegment ? borderCell.ColumnSpan : 1;
                                                    double borderHeight = rowHeight;
                                                    double borderBottom = rowBottom;
                                                    if (wholeRowSegment) {
                                                        if (borderCell.RowSpan > 1) {
                                                            borderHeight = GetTableCellHeight(table.RowHeights, rowIndex, borderCell.RowSpan);
                                                            borderBottom = yCol - borderHeight;
                                                        }
                                                    }

                                                    pageDirty = true;
                                                    DrawCellBorder(sb, cellBorder, borderX, borderBottom, GetTableCellWidth(table.ColumnWidths, borderColumn, span, 0D), borderHeight);
                                                }
                                                borderX += table.ColumnWidths[borderColumn];
                                            }
                                        }

                                        yCol -= rowHeight;
                                        remain -= rowHeight;
                                        consumed += rowHeight;
                                    }

                                    void DrawColumnTableRow(int rowIndex, bool renderAsHeader) =>
                                        DrawColumnTableRowSegment(rowIndex, renderAsHeader, 0, table.RowLineCounts[rowIndex]);

                                    int rowIndex = line;
                                    int rowStartLine = subline;
                                    while (rowIndex < tbColumn.Rows.Count) {
                                        double rowHeight = table.RowHeights[rowIndex];
                                        if (rowHeight > maxContentHeight + 0.001) {
                                            if (!tableStyle.AllowRowBreakAcrossPages) {
                                                throw new ArgumentException("Table row height exceeds the available page content height and row splitting is disabled.");
                                            }

                                            int totalLines = table.RowLineCounts[rowIndex];
                                            bool repeatHeaderBeforeSegment = rowIndex >= table.HeaderRowCount &&
                                                HasRepeatableHeader() &&
                                                AtContinuationPageTop() &&
                                                repeatHeaderHeight + table.RowLeadings[rowIndex] + padTop + padBottom <= remain + 0.001;
                                            double neededForFirstSegment = table.RowLeadings[rowIndex] + padTop + padBottom + (repeatHeaderBeforeSegment ? repeatHeaderHeight : 0);
                                            if (neededForFirstSegment > remain && consumed > 0) break;
                                            if (neededForFirstSegment > remain && consumed == 0) { remain = 0; break; }

                                            if (repeatHeaderBeforeSegment) {
                                                for (int headerIndex = 0; headerIndex < table.HeaderRowCount; headerIndex++) {
                                                    DrawColumnTableRow(headerIndex, renderAsHeader: true);
                                                }
                                            }

                                            int maxLinesThisPage = Math.Max(1, (int)Math.Floor((remain - padTop - padBottom) / table.RowLeadings[rowIndex]));
                                            int take = Math.Min(totalLines - rowStartLine, maxLinesThisPage);
                                            DrawColumnTableRowSegment(rowIndex, renderAsHeader: rowIndex < table.HeaderRowCount && rowStartLine == 0, rowStartLine, take);
                                            rowStartLine += take;

                                            if (rowStartLine < totalLines) {
                                                line = rowIndex;
                                                subline = rowStartLine;
                                                break;
                                            }

                                            rowIndex++;
                                            line = rowIndex;
                                            subline = 0;
                                            rowStartLine = 0;
                                            continue;
                                        }
                                        bool repeatHeaderBeforeRow = rowIndex >= table.HeaderRowCount &&
                                            HasRepeatableHeader() &&
                                            AtContinuationPageTop() &&
                                            repeatHeaderHeight + rowHeight <= remain + 0.001;
                                        double neededForNextRow = rowHeight + (repeatHeaderBeforeRow ? repeatHeaderHeight : 0);
                                        if (neededForNextRow > remain && consumed > 0) break;
                                        if (neededForNextRow > remain && consumed == 0) { remain = 0; break; }

                                        if (repeatHeaderBeforeRow) {
                                            for (int headerIndex = 0; headerIndex < table.HeaderRowCount; headerIndex++) {
                                                DrawColumnTableRow(headerIndex, renderAsHeader: true);
                                            }
                                        }

                                        DrawColumnTableRow(rowIndex, renderAsHeader: rowIndex < table.HeaderRowCount);
                                        rowIndex++;
                                        line = rowIndex;
                                        subline = 0;
                                        rowStartLine = 0;
                                    }

                                    if (rowIndex >= tbColumn.Rows.Count) {
                                        if (tableStyle.SpacingAfter > 0 && tableStyle.SpacingAfter <= remain) {
                                            yCol -= tableStyle.SpacingAfter;
                                            remain -= tableStyle.SpacingAfter;
                                            consumed += tableStyle.SpacingAfter;
                                        }
                                        idx++;
                                        line = 0;
                                        subline = 0;
                                    } else {
                                        break;
                                    }
                                } else if (it is ColRule cr) {
                                    PdfHorizontalRuleStyle hr2 = ResolveHorizontalRuleStyle(cr.Block, currentOpts);
                                    ValidateHorizontalRule(hr2);
                                    double needed = hr2.SpacingBefore + hr2.Thickness + hr2.SpacingAfter;
                                    EnsureFixedFlowBlockFits("Horizontal rule", wCol, needed, wCol);
                                    if (line == 0 && hr2.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = needed + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    yCol -= hr2.SpacingBefore;
                                    double x1 = xCol, x2 = xCol + wCol, yLine = yCol - hr2.Thickness * 0.5;
                                    pageDirty = true;
                                    DrawHLine(sb, hr2.Color, hr2.Thickness, x1, x2, yLine);
                                    yCol -= hr2.Thickness + hr2.SpacingAfter; remain -= needed; consumed += needed; idx++;
                                } else if (it is ColImg ciimg) {
                                    var ib2 = ciimg.Block;
                                    PdfImageStyle imageStyle = ResolveImageStyle(ib2, currentOpts);
                                    PdfDoc.ValidateImageStyleForBox(imageStyle, ib2.Width, ib2.Height, nameof(imageStyle.ClipPath));
                                    PdfDoc.ValidateImageFitDimensions(ib2.Info, imageStyle.Fit, nameof(imageStyle.Fit));
                                    double needed = imageStyle.SpacingBefore + ib2.Height + imageStyle.SpacingAfter;
                                    EnsureFixedFlowBlockFits("Image", ib2.Width, needed, wCol);
                                    if (imageStyle.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = needed + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    yCol -= imageStyle.SpacingBefore;
                                    double xImg = xCol;
                                    if (imageStyle.Align == PdfAlign.Center) xImg = xCol + Math.Max(0, (wCol - ib2.Width) / 2);
                                    else if (imageStyle.Align == PdfAlign.Right) xImg = xCol + Math.Max(0, wCol - ib2.Width);
                                    PageImage pageImage = CreatePageImage(ib2, imageStyle, xImg, yCol - ib2.Height);
                                    currentPage!.Images.Add(pageImage);
                                    AddImageLinkAnnotation(ib2, imageStyle, pageImage, xImg, yCol - ib2.Height);
                                    pageDirty = true;
                                    yCol -= ib2.Height + imageStyle.SpacingAfter; remain -= needed; consumed += needed; idx++;
                                } else if (it is ColShape cs) {
                                    var shape = cs.Block;
                                    PdfDrawingStyle shapeStyle = ResolveDrawingStyle(shape, currentOpts);
                                    PdfDoc.ValidateDrawingStyle(shapeStyle, "Shape");
                                    double needed = shapeStyle.SpacingBefore + shape.Shape.Height + shapeStyle.SpacingAfter;
                                    EnsureFixedFlowBlockFits("Shape", shape.Shape.Width, needed, wCol);
                                    if (shapeStyle.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = needed + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    yCol -= shapeStyle.SpacingBefore;
                                    DrawShapeAt(shape, shapeStyle, xCol, wCol, yCol);
                                    AddShapeLinkAnnotation(shape, shapeStyle, xCol, wCol, yCol);
                                    yCol -= shape.Shape.Height + shapeStyle.SpacingAfter;
                                    remain -= needed;
                                    consumed += needed;
                                    idx++;
                                } else if (it is ColDrawing cd) {
                                    var drawing = cd.Block;
                                    PdfDrawingStyle drawingStyle = ResolveDrawingStyle(drawing, currentOpts);
                                    PdfDoc.ValidateDrawingStyle(drawingStyle, "Drawing");
                                    double needed = drawingStyle.SpacingBefore + drawing.Drawing.Height + drawingStyle.SpacingAfter;
                                    EnsureFixedFlowBlockFits("Drawing", drawing.Drawing.Width, needed, wCol);
                                    if (drawingStyle.KeepWithNext && idx + 1 < items.Count) {
                                        double nextHeight = MeasureColItemFirstVisualHeight(items[idx + 1]);
                                        double keepHeight = needed + nextHeight;
                                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && keepHeight > remain + 0.001) {
                                            if (consumed > 0) break;
                                            remain = 0;
                                            break;
                                        }
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    yCol -= drawingStyle.SpacingBefore;
                                    DrawDrawingAt(drawing, drawingStyle, xCol, wCol, yCol);
                                    AddDrawingLinkAnnotation(drawing, drawingStyle, xCol, wCol, yCol);
                                    yCol -= drawing.Drawing.Height + drawingStyle.SpacingAfter;
                                    remain -= needed;
                                    consumed += needed;
                                    idx++;
                                } else if (it is ColBookmark bookmarkItem) {
                                    AddNamedDestination(bookmarkItem.Block, yCol);
                                    idx++;
                                } else if (it is ColSpacer spacerItem) {
                                    double needed = spacerItem.Block.Height;
                                    double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                                    if (needed > availableHeight + 0.001) {
                                        throw new ArgumentException("Spacer height exceeds the available page content height.");
                                    }

                                    if (needed > remain && consumed > 0) break;
                                    if (needed > remain && consumed == 0) { remain = 0; break; }
                                    yCol -= needed;
                                    remain -= needed;
                                    consumed += needed;
                                    idx++;
                                }
                            }
                            colStates[ci] = (idx, line, subline);
                            if (colStates[ci] != startState) {
                                anyColumnAdvanced = true;
                            }

                            if (consumed > maxConsumed) maxConsumed = consumed;
                        }

                        if (maxConsumed <= 0.01) {
                            if (anyColumnAdvanced && !AnyRemaining()) {
                                break;
                            }

                            if (Math.Abs(y - yStart) <= 0.001) {
                                throw new InvalidOperationException("Row column layout could not make progress on an empty page.");
                            }

                            NewPage();
                            continue;
                        }
                        y -= maxConsumed;
                    }

                    if (rowSpacingAfter > 0) {
                        y -= rowSpacingAfter;
                    }
                } else if (block is ImageBlock ib) {
                    double xImg = currentOpts.MarginLeft;
                    double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                    PdfImageStyle imageStyle = ResolveImageStyle(ib, currentOpts);
                    PdfDoc.ValidateImageStyleForBox(imageStyle, ib.Width, ib.Height, nameof(imageStyle.ClipPath));
                    PdfDoc.ValidateImageFitDimensions(ib.Info, imageStyle.Fit, nameof(imageStyle.Fit));
                    double needed = imageStyle.SpacingBefore + ib.Height + imageStyle.SpacingAfter;
                    if (imageStyle.Align == PdfAlign.Center) xImg = currentOpts.MarginLeft + Math.Max(0, (contentWidth - ib.Width) / 2);
                    else if (imageStyle.Align == PdfAlign.Right) xImg = currentOpts.MarginLeft + Math.Max(0, contentWidth - ib.Width);
                    EnsureFixedFlowBlockFits("Image", ib.Width, needed, contentWidth);
                    if (imageStyle.KeepWithNext && nextBlock != null) {
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                        double keepHeight = needed + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    if (y - needed < currentOpts.MarginBottom) { NewPage(); }
                    y -= imageStyle.SpacingBefore;
                    EnsurePage();
                    PageImage pageImage = CreatePageImage(ib, imageStyle, xImg, y - ib.Height);
                    currentPage!.Images.Add(pageImage);
                    AddImageLinkAnnotation(ib, imageStyle, pageImage, xImg, y - ib.Height);
                    pageDirty = true;
                    y -= ib.Height + imageStyle.SpacingAfter;
                } else if (block is PanelParagraphBlock ppb) {
                    double size = currentOpts.DefaultFontSize;
                    double leading = size * 1.4;
                    var panelFont = ChooseNormal(currentOpts.DefaultFont);
                    double firstBaselineOffset = GetAscender(panelFont, size);
                    double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                    PanelStyle panelStyle = ResolvePanelStyle(ppb, currentOpts);
                    double innerWidth = panelStyle.MaxWidth.HasValue ? Math.Min(contentWidth, panelStyle.MaxWidth.Value) : contentWidth;
                    ValidatePanelStyle(panelStyle, innerWidth);
                    double textWidthAvail = innerWidth - 2 * panelStyle.PaddingX;
                    var (lines, lineHeights) = WrapRichRuns(ppb.Runs, textWidthAvail, size, panelFont, leading);
                    double panelWidth = innerWidth;
                    double xLeft = currentOpts.MarginLeft;
                    if (panelStyle.Align == PdfAlign.Center) xLeft = currentOpts.MarginLeft + Math.Max(0, (contentWidth - innerWidth) / 2);
                    else if (panelStyle.Align == PdfAlign.Right) xLeft = currentOpts.MarginLeft + Math.Max(0, contentWidth - innerWidth);

                    if (panelStyle.KeepWithNext && nextBlock != null && lines.Count > 0) {
                        double panelHeight = panelStyle.SpacingBefore + panelStyle.PaddingY + lineHeights.Sum() + panelStyle.PaddingY + panelStyle.SpacingAfter;
                        double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                        double keepHeight = panelHeight + nextHeight;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                            NewPage();
                        }
                    }

                    if (panelStyle.SpacingBefore > 0) {
                        if (y < yStart - 0.001 && y - panelStyle.SpacingBefore < currentOpts.MarginBottom) {
                            NewPage();
                        }

                        y -= panelStyle.SpacingBefore;
                    }

                    if (panelStyle.KeepTogether) {
                        double textHeight = lineHeights.Sum();
                        double panelHeight = panelStyle.PaddingY + textHeight + panelStyle.PaddingY;
                        double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                        if (panelHeight > availableHeight + 0.001) {
                            throw new ArgumentException("Panel height exceeds the available page content height.");
                        }

                        double panelTop = y;
                        double panelBottom = y - panelHeight;
                        if (panelBottom < currentOpts.MarginBottom) { NewPage(); panelTop = y; panelBottom = y - panelHeight; }
                        if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xLeft, panelBottom, panelWidth, panelTop - panelBottom); }
                        if (panelStyle.BorderColor.HasValue && panelStyle.BorderWidth > 0) { pageDirty = true; DrawRowRect(sb, panelStyle.BorderColor.Value, panelStyle.BorderWidth, xLeft, panelBottom, panelWidth, panelTop - panelBottom); }
                        pageDirty = true;
                        WriteRichParagraph(sb, new RichParagraphBlock(ppb.Runs, ppb.Align, ppb.DefaultColor), lines, lineHeights, currentOpts, panelTop - panelStyle.PaddingY - firstBaselineOffset, size, leading, currentPage!.Annotations, xLeft + panelStyle.PaddingX, textWidthAvail);
                        if (ppb.Runs.Any(r => r.Bold)) { currentPage!.UsedBold = true; usedBold = true; }
                        if (ppb.Runs.Any(r => r.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
                        if (ppb.Runs.Any(r => r.Bold && r.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }
                        y = panelBottom;
                        if (panelStyle.SpacingAfter > 0) {
                            if (y < yStart - 0.001 && y - panelStyle.SpacingAfter < currentOpts.MarginBottom) {
                                NewPage();
                            } else {
                                y -= panelStyle.SpacingAfter;
                            }
                        }
                    } else {
                        int li = 0; bool firstSeg = true;
                        while (li < lines.Count) {
                            double avail = y - currentOpts.MarginBottom;
                            if (avail < 0.5) { NewPage(); firstSeg = false; continue; }
                            double topPad = firstSeg ? panelStyle.PaddingY : 0;
                            double minLine = lineHeights[li];
                            if (avail < topPad + minLine) { NewPage(); firstSeg = false; continue; }
                            double roomForText = avail - topPad - panelStyle.PaddingY;
                            int take = 0; double hsum = 0;
                            for (int k = li; k < lines.Count; k++) {
                                double h = lineHeights[k];
                                if (hsum + h > roomForText) break;
                                hsum += h; take++;
                            }
                            bool lastSeg = (li + take) >= lines.Count;
                            double panelTop = y;
                            double usedBottomPad = panelStyle.PaddingY;
                            if (!lastSeg && topPad + hsum + usedBottomPad > avail) usedBottomPad = Math.Max(0, avail - (topPad + hsum));
                            double panelBottom = y - (topPad + hsum + usedBottomPad);
                            if (panelStyle.Background.HasValue) { pageDirty = true; DrawRowFill(sb, panelStyle.Background.Value, xLeft, panelBottom, panelWidth, panelTop - panelBottom); }
                            if (panelStyle.BorderColor.HasValue && panelStyle.BorderWidth > 0) { pageDirty = true; DrawRowRect(sb, panelStyle.BorderColor.Value, panelStyle.BorderWidth, xLeft, panelBottom, panelWidth, panelTop - panelBottom); }
                            var sliceLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
                            var sliceHeights = new System.Collections.Generic.List<double>();
                            for (int k = 0; k < take; k++) { sliceLines.Add(lines[li + k]); sliceHeights.Add(lineHeights[li + k]); }
                            pageDirty = true;
                            WriteRichParagraph(sb, new RichParagraphBlock(ppb.Runs, ppb.Align, ppb.DefaultColor), sliceLines, sliceHeights, currentOpts, panelTop - topPad - firstBaselineOffset, size, leading, currentPage!.Annotations, xLeft + panelStyle.PaddingX, textWidthAvail);
                            if (ppb.Runs.Any(r => r.Bold)) { currentPage!.UsedBold = true; usedBold = true; }
                            if (ppb.Runs.Any(r => r.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
                            if (ppb.Runs.Any(r => r.Bold && r.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }
                            y = panelBottom; li += take; firstSeg = false;
                            if (li < lines.Count) {
                                NewPage();
                            } else if (panelStyle.SpacingAfter > 0) {
                                if (y < yStart - 0.001 && y - panelStyle.SpacingAfter < currentOpts.MarginBottom) {
                                    NewPage();
                                } else {
                                    y -= panelStyle.SpacingAfter;
                                }
                            }
                        }
                    }
                }
            }
        }

        ProcessBlocks(blocks);
        FlushPage(pageDirty || (currentPage?.Images.Count ?? 0) > 0 || (currentPage?.Annotations.Count ?? 0) > 0 || (currentPage?.GraphicsStates.Count ?? 0) > 0 || (currentPage?.Shadings.Count ?? 0) > 0 || (currentPage?.NamedDestinations.Count ?? 0) > 0);

        var result = new LayoutResult { UsedBold = usedBold, UsedItalic = usedItalic, UsedBoldItalic = usedBoldItalic };
        foreach (var p in pages) result.Pages.Add(p);
        return result;
    }

}
