using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendTable(StringBuilder text, List<LegacyDocWritableRun> runs, List<LegacyDocWritableParagraph> paragraphFormats, Table table) {
            ThrowIfUnsupportedTableShape(table);

            TableRow[] rows = table.Elements<TableRow>().ToArray();
            if (rows.Length == 0) {
                throw new NotSupportedException("Native DOC saving supports simple tables only when at least one row is present.");
            }

            IReadOnlyList<int> gridColumnWidthsTwips = ReadSupportedTableGridWidths(table.GetFirstChild<TableGrid>());
            foreach (TableRow row in rows) {
                LegacyDocWritableTableRowFormatting rowFormatting = ReadSupportedTableRowFormatting(row, out TableCell[] cells);
                if (cells.Length == 0) {
                    throw new NotSupportedException("Native DOC saving supports simple tables only when every row contains at least one cell.");
                }

                IReadOnlyList<LegacyDocWritableTableCell> writableCells = ExpandSupportedTableCells(cells, gridColumnWidthsTwips);
                IReadOnlyList<int> cellWidthsTwips = ReadSupportedTableCellWidths(writableCells);
                IReadOnlyList<LegacyDocTableCellHorizontalMerge> cellHorizontalMerges = ReadSupportedTableCellHorizontalMerges(writableCells);
                IReadOnlyList<LegacyDocTableCellVerticalMerge> cellVerticalMerges = ReadSupportedTableCellVerticalMerges(writableCells);
                foreach (LegacyDocWritableTableCell writableCell in writableCells) {
                    int cellStart = text.Length;
                    LegacyDocWritableParagraphFormatting paragraphFormatting = AppendTableCell(text, runs, writableCell.SourceCell)
                        .WithTableMarkers(isTableTerminatingParagraph: false);
                    text.Append('\a');
                    paragraphFormats.Add(new LegacyDocWritableParagraph(cellStart, text.Length - cellStart, paragraphFormatting));
                }

                int rowTerminatorStart = text.Length;
                text.Append('\a');
                paragraphFormats.Add(new LegacyDocWritableParagraph(
                    rowTerminatorStart,
                    1,
                    LegacyDocWritableParagraphFormatting.Plain.WithTableMarkers(
                        isTableTerminatingParagraph: true,
                        tableCellWidthsTwips: cellWidthsTwips,
                        tableRowHeightTwips: rowFormatting.RowHeightTwips,
                        tableRowHeightIsExact: rowFormatting.RowHeightIsExact,
                        tableRowCantSplit: rowFormatting.RowCantSplit,
                        tableRowIsHeader: rowFormatting.RowIsHeader,
                        tableCellHorizontalMerges: cellHorizontalMerges,
                        tableCellVerticalMerges: cellVerticalMerges)));
            }

            text.Append('\r');
        }

        private static void ThrowIfUnsupportedTableShape(Table table) {
            foreach (OpenXmlElement child in table.ChildElements) {
                switch (child) {
                    case TableProperties tableProperties:
                        ThrowIfUnsupportedTableProperties(tableProperties);
                        break;
                    case TableGrid tableGrid:
                        ThrowIfUnsupportedTableGrid(tableGrid);
                        break;
                    case TableRow:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table element: {child.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableProperties(TableProperties tableProperties) {
            foreach (OpenXmlElement property in tableProperties.ChildElements) {
                switch (property) {
                    case TableStyle tableStyle:
                        if (!string.Equals(tableStyle.Val?.Value, "TableGrid", StringComparison.OrdinalIgnoreCase)) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with the TableGrid table style.");
                        }
                        break;
                    case TableWidth tableWidth:
                        if (tableWidth.Type?.Value != TableWidthUnitValues.Auto || tableWidth.Width?.Value != "0") {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with the default automatic table width.");
                        }
                        break;
                    case TableLook:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table property: {property.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableGrid(TableGrid tableGrid) {
            foreach (OpenXmlElement child in tableGrid.ChildElements) {
                if (child is GridColumn) {
                    continue;
                }

                throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table grid element: {child.LocalName}.");
            }
        }

        private static IReadOnlyList<int> ReadSupportedTableGridWidths(TableGrid? tableGrid) {
            if (tableGrid == null) {
                return Array.Empty<int>();
            }

            GridColumn[] columns = tableGrid.Elements<GridColumn>().ToArray();
            var widths = new int[columns.Length];
            for (int index = 0; index < columns.Length; index++) {
                string? widthText = columns[index].Width?.Value;
                if (string.IsNullOrWhiteSpace(widthText)) {
                    widths[index] = 0;
                    continue;
                }

                if (!int.TryParse(widthText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int width)
                    || width <= 0
                    || width > short.MaxValue) {
                    throw new NotSupportedException("Native DOC saving supports table grid column widths only as positive DXA twip values within the Word 97-2003 signed twip range.");
                }

                widths[index] = width;
            }

            return widths;
        }

        private static LegacyDocWritableTableRowFormatting ReadSupportedTableRowFormatting(TableRow row, out TableCell[] cells) {
            var tableCells = new List<TableCell>();
            LegacyDocWritableTableRowFormatting rowFormatting = LegacyDocWritableTableRowFormatting.Empty;
            bool hasTableRowProperties = false;
            foreach (OpenXmlElement child in row.ChildElements) {
                switch (child) {
                    case TableRowProperties tableRowProperties:
                        if (hasTableRowProperties) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one table row property collection per row.");
                        }

                        rowFormatting = ReadSupportedTableRowProperties(tableRowProperties);
                        hasTableRowProperties = true;
                        break;
                    case TableCell tableCell:
                        tableCells.Add(tableCell);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table row element: {child.LocalName}.");
                }
            }

            cells = tableCells.ToArray();
            return rowFormatting;
        }

        private static LegacyDocWritableTableRowFormatting ReadSupportedTableRowProperties(TableRowProperties tableRowProperties) {
            int? rowHeightTwips = null;
            bool rowHeightIsExact = false;
            bool? rowCantSplit = null;
            bool? rowIsHeader = null;
            bool hasRowHeight = false;
            bool hasCantSplit = false;
            bool hasTableHeader = false;
            foreach (OpenXmlElement property in tableRowProperties.ChildElements) {
                switch (property) {
                    case TableRowHeight rowHeight:
                        if (hasRowHeight) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one table row height per row.");
                        }

                        ReadSupportedTableRowHeight(rowHeight, out rowHeightTwips, out rowHeightIsExact);
                        hasRowHeight = true;
                        break;
                    case CantSplit cantSplit:
                        if (hasCantSplit) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one row no-split flag per row.");
                        }

                        rowCantSplit = ReadTableRowOnOff(cantSplit);
                        hasCantSplit = true;
                        break;
                    case TableHeader tableHeader:
                        if (hasTableHeader) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one row header flag per row.");
                        }

                        rowIsHeader = ReadTableRowOnOff(tableHeader);
                        hasTableHeader = true;
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table row property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableTableRowFormatting(rowHeightTwips, rowHeightIsExact, rowCantSplit, rowIsHeader);
        }

        private static void ReadSupportedTableRowHeight(TableRowHeight rowHeight, out int? rowHeightTwips, out bool rowHeightIsExact) {
            rowHeightTwips = null;
            rowHeightIsExact = false;

            uint? rawValue = rowHeight.Val?.Value;
            if (rawValue == null || rawValue.Value == 0) {
                return;
            }

            if (rawValue.Value > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table row heights only as positive twip values within the Word 97-2003 signed twip range.");
            }

            HeightRuleValues? heightRule = rowHeight.HeightType?.Value;
            if (heightRule == HeightRuleValues.Auto) {
                throw new NotSupportedException("Native DOC saving supports table row heights only with exact or at-least height rules.");
            }

            if (heightRule != null && heightRule != HeightRuleValues.Exact && heightRule != HeightRuleValues.AtLeast) {
                throw new NotSupportedException($"Native DOC saving does not support table row height rule '{heightRule}'.");
            }

            rowHeightTwips = checked((int)rawValue.Value);
            rowHeightIsExact = heightRule == HeightRuleValues.Exact;
        }

        private static bool? ReadTableRowOnOff(CantSplit cantSplit) {
            if (cantSplit.Val == null) {
                return true;
            }

            return ReadTableRowOnOffValue(cantSplit.Val.InnerText) ? true : null;
        }

        private static bool? ReadTableRowOnOff(TableHeader tableHeader) {
            if (tableHeader.Val == null) {
                return true;
            }

            return ReadTableRowOnOffValue(tableHeader.Val.InnerText) ? true : null;
        }

        private static bool ReadTableRowOnOffValue(string? value) =>
            !string.Equals(value, "0", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(value, "off", StringComparison.OrdinalIgnoreCase);

        private static IReadOnlyList<LegacyDocWritableTableCell> ExpandSupportedTableCells(IReadOnlyList<TableCell> cells, IReadOnlyList<int> gridColumnWidthsTwips) {
            var writableCells = new List<LegacyDocWritableTableCell>();
            int logicalColumnIndex = 0;
            foreach (TableCell cell in cells) {
                TableCellProperties? cellProperties = cell.TableCellProperties;
                int gridSpan = ReadSupportedGridSpan(cellProperties);
                LegacyDocTableCellHorizontalMerge horizontalMerge = ReadSupportedTableCellHorizontalMerge(cell);
                LegacyDocTableCellVerticalMerge verticalMerge = ReadSupportedTableCellVerticalMerge(cell);
                if (gridSpan > 1 && horizontalMerge == LegacyDocTableCellHorizontalMerge.Continue) {
                    throw new NotSupportedException("Native DOC saving supports simple horizontal table cell merges only. A continued horizontal merge cannot also define gridSpan.");
                }

                for (int spanIndex = 0; spanIndex < gridSpan; spanIndex++) {
                    int width = ReadSupportedTableCellWidthForSpan(cellProperties, gridColumnWidthsTwips, logicalColumnIndex, spanIndex, gridSpan);
                    LegacyDocTableCellHorizontalMerge merge = gridSpan == 1
                        ? horizontalMerge
                        : spanIndex == 0
                            ? LegacyDocTableCellHorizontalMerge.Restart
                            : LegacyDocTableCellHorizontalMerge.Continue;
                    writableCells.Add(new LegacyDocWritableTableCell(spanIndex == 0 ? cell : null, width, merge, verticalMerge));
                }

                logicalColumnIndex += gridSpan;
            }

            return writableCells;
        }

        private static IReadOnlyList<int> ReadSupportedTableCellWidths(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var widths = new int[cells.Count];
            for (int index = 0; index < cells.Count; index++) {
                widths[index] = cells[index].WidthTwips;
            }

            return widths;
        }

        private static IReadOnlyList<LegacyDocTableCellHorizontalMerge> ReadSupportedTableCellHorizontalMerges(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var merges = new LegacyDocTableCellHorizontalMerge[cells.Count];
            bool hasMerge = false;
            for (int index = 0; index < cells.Count; index++) {
                merges[index] = cells[index].HorizontalMerge;
                if (merges[index] != LegacyDocTableCellHorizontalMerge.None) {
                    hasMerge = true;
                }
            }

            return hasMerge ? merges : Array.Empty<LegacyDocTableCellHorizontalMerge>();
        }

        private static IReadOnlyList<LegacyDocTableCellVerticalMerge> ReadSupportedTableCellVerticalMerges(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var merges = new LegacyDocTableCellVerticalMerge[cells.Count];
            bool hasMerge = false;
            for (int index = 0; index < cells.Count; index++) {
                merges[index] = cells[index].VerticalMerge;
                if (merges[index] != LegacyDocTableCellVerticalMerge.None) {
                    hasMerge = true;
                }
            }

            return hasMerge ? merges : Array.Empty<LegacyDocTableCellVerticalMerge>();
        }

        private static LegacyDocTableCellHorizontalMerge ReadSupportedTableCellHorizontalMerge(TableCell cell) {
            HorizontalMerge? horizontalMerge = cell.TableCellProperties?.GetFirstChild<HorizontalMerge>();
            if (horizontalMerge == null) {
                return LegacyDocTableCellHorizontalMerge.None;
            }

            MergedCellValues? value = horizontalMerge.Val?.Value;
            if (value == MergedCellValues.Restart) {
                return LegacyDocTableCellHorizontalMerge.Restart;
            }

            if (value == MergedCellValues.Continue) {
                return LegacyDocTableCellHorizontalMerge.Continue;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell horizontal merge value '{value}'.");
        }

        private static LegacyDocTableCellVerticalMerge ReadSupportedTableCellVerticalMerge(TableCell cell) {
            VerticalMerge? verticalMerge = cell.TableCellProperties?.GetFirstChild<VerticalMerge>();
            if (verticalMerge == null) {
                return LegacyDocTableCellVerticalMerge.None;
            }

            MergedCellValues? value = verticalMerge.Val?.Value;
            if (value == MergedCellValues.Restart) {
                return LegacyDocTableCellVerticalMerge.Restart;
            }

            if (value == MergedCellValues.Continue) {
                return LegacyDocTableCellVerticalMerge.Continue;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell vertical merge value '{value}'.");
        }

        private static int GetGridColumnWidth(IReadOnlyList<int> gridColumnWidthsTwips, int columnIndex) {
            return columnIndex < gridColumnWidthsTwips.Count ? gridColumnWidthsTwips[columnIndex] : 0;
        }

        private static int ReadSupportedTableCellWidth(TableCellProperties? cellProperties, int gridColumnWidthTwips) {
            int? explicitWidth = ReadSupportedExplicitTableCellWidth(cellProperties);
            if (explicitWidth != null) {
                return explicitWidth.Value;
            }

            return gridColumnWidthTwips > 0 ? gridColumnWidthTwips : 2400;
        }

        private static int ReadSupportedTableCellWidthForSpan(TableCellProperties? cellProperties, IReadOnlyList<int> gridColumnWidthsTwips, int logicalColumnIndex, int spanIndex, int gridSpan) {
            int gridColumnWidth = GetGridColumnWidth(gridColumnWidthsTwips, logicalColumnIndex + spanIndex);
            if (gridColumnWidth > 0) {
                return gridColumnWidth;
            }

            int? explicitWidth = ReadSupportedExplicitTableCellWidth(cellProperties);
            if (explicitWidth == null) {
                return 2400;
            }

            if (gridSpan == 1) {
                return explicitWidth.Value;
            }

            int baseWidth = explicitWidth.Value / gridSpan;
            int remainder = explicitWidth.Value % gridSpan;
            int width = baseWidth + (spanIndex < remainder ? 1 : 0);
            return width > 0 ? width : 1;
        }

        private static int? ReadSupportedExplicitTableCellWidth(TableCellProperties? cellProperties) {
            TableCellWidth? cellWidth = cellProperties?.GetFirstChild<TableCellWidth>();
            if (cellWidth == null) {
                return null;
            }

            if (cellWidth.Type?.Value != TableWidthUnitValues.Dxa) {
                throw new NotSupportedException("Native DOC saving supports simple table cell widths only as explicit DXA twip values.");
            }

            string? widthText = cellWidth.Width?.Value;
            if (string.IsNullOrWhiteSpace(widthText)
                || !int.TryParse(widthText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int width)
                || width <= 0
                || width > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports simple table cell widths only within the Word 97-2003 signed twip range.");
            }

            return width;
        }

        private static int ReadSupportedGridSpan(TableCellProperties? cellProperties) {
            GridSpan? gridSpan = cellProperties?.GetFirstChild<GridSpan>();
            if (gridSpan == null) {
                return 1;
            }

            int span = gridSpan.Val?.Value ?? 1;
            if (span <= 0 || span > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table cell gridSpan only as a positive value within the DOC table column limit.");
            }

            return span;
        }

        private static LegacyDocWritableParagraphFormatting AppendTableCell(StringBuilder text, List<LegacyDocWritableRun> runs, TableCell? cell) {
            if (cell == null) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            if (cell.Elements<Table>().Any()) {
                throw new NotSupportedException("Native DOC saving supports simple tables only. Nested tables are not supported yet.");
            }

            Paragraph? paragraph = null;
            foreach (OpenXmlElement child in cell.ChildElements) {
                switch (child) {
                    case TableCellProperties cellProperties:
                        ThrowIfUnsupportedTableCellProperties(cellProperties);
                        break;
                    case Paragraph cellParagraph:
                        if (paragraph != null) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one paragraph per cell.");
                        }

                        paragraph = cellParagraph;
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table cell element: {child.LocalName}.");
                }
            }

            if (paragraph != null) {
                return AppendTableCellParagraph(text, runs, paragraph);
            }

            return LegacyDocWritableParagraphFormatting.Plain;
        }

        private static void ThrowIfUnsupportedTableCellProperties(TableCellProperties cellProperties) {
            foreach (OpenXmlElement property in cellProperties.ChildElements) {
                switch (property) {
                    case TableCellWidth cellWidth:
                        ReadSupportedTableCellWidth(cellProperties, gridColumnWidthTwips: 0);
                        break;
                    case GridSpan:
                        ReadSupportedGridSpan(cellProperties);
                        break;
                    case HorizontalMerge:
                        break;
                    case VerticalMerge:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table cell property: {property.LocalName}.");
                }
            }
        }

        private static LegacyDocWritableParagraphFormatting AppendTableCellParagraph(StringBuilder text, List<LegacyDocWritableRun> runs, Paragraph paragraph) {
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedParagraphFormatting(paragraph.ParagraphProperties);

            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        AppendSupportedRunText(text, runs, run);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple table cell paragraphs only with text runs. Unsupported paragraph element: {child.LocalName}.");
                }
            }

            return paragraphFormatting;
        }

        private readonly struct LegacyDocWritableTableRowFormatting {
            internal static readonly LegacyDocWritableTableRowFormatting Empty = new LegacyDocWritableTableRowFormatting(null, false, null, null);

            internal LegacyDocWritableTableRowFormatting(int? rowHeightTwips, bool rowHeightIsExact, bool? rowCantSplit, bool? rowIsHeader) {
                RowHeightTwips = rowHeightTwips;
                RowHeightIsExact = rowHeightIsExact;
                RowCantSplit = rowCantSplit;
                RowIsHeader = rowIsHeader;
            }

            internal int? RowHeightTwips { get; }

            internal bool RowHeightIsExact { get; }

            internal bool? RowCantSplit { get; }

            internal bool? RowIsHeader { get; }
        }

        private readonly struct LegacyDocWritableTableCell {
            internal LegacyDocWritableTableCell(TableCell? sourceCell, int widthTwips, LegacyDocTableCellHorizontalMerge horizontalMerge, LegacyDocTableCellVerticalMerge verticalMerge) {
                SourceCell = sourceCell;
                WidthTwips = widthTwips;
                HorizontalMerge = horizontalMerge;
                VerticalMerge = verticalMerge;
            }

            internal TableCell? SourceCell { get; }

            internal int WidthTwips { get; }

            internal LegacyDocTableCellHorizontalMerge HorizontalMerge { get; }

            internal LegacyDocTableCellVerticalMerge VerticalMerge { get; }
        }
    }
}
