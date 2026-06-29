using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
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

                IReadOnlyList<int> cellWidthsTwips = ReadSupportedTableCellWidths(cells, gridColumnWidthsTwips);
                foreach (TableCell cell in cells) {
                    int cellStart = text.Length;
                    LegacyDocWritableParagraphFormatting paragraphFormatting = AppendTableCell(text, runs, cell)
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
                        cellWidthsTwips,
                        rowFormatting.RowHeightTwips,
                        rowFormatting.RowHeightIsExact)));
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
            bool hasRowHeight = false;
            foreach (OpenXmlElement property in tableRowProperties.ChildElements) {
                switch (property) {
                    case TableRowHeight rowHeight:
                        if (hasRowHeight) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one table row height per row.");
                        }

                        ReadSupportedTableRowHeight(rowHeight, out rowHeightTwips, out rowHeightIsExact);
                        hasRowHeight = true;
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table row property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableTableRowFormatting(rowHeightTwips, rowHeightIsExact);
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

        private static IReadOnlyList<int> ReadSupportedTableCellWidths(IReadOnlyList<TableCell> cells, IReadOnlyList<int> gridColumnWidthsTwips) {
            var widths = new int[cells.Count];
            for (int index = 0; index < cells.Count; index++) {
                widths[index] = ReadSupportedTableCellWidth(cells[index].TableCellProperties, GetGridColumnWidth(gridColumnWidthsTwips, index));
            }

            return widths;
        }

        private static int GetGridColumnWidth(IReadOnlyList<int> gridColumnWidthsTwips, int columnIndex) {
            return columnIndex < gridColumnWidthsTwips.Count ? gridColumnWidthsTwips[columnIndex] : 0;
        }

        private static int ReadSupportedTableCellWidth(TableCellProperties? cellProperties, int gridColumnWidthTwips) {
            TableCellWidth? cellWidth = cellProperties?.GetFirstChild<TableCellWidth>();
            if (cellWidth == null) {
                return gridColumnWidthTwips > 0 ? gridColumnWidthTwips : 2400;
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

        private static LegacyDocWritableParagraphFormatting AppendTableCell(StringBuilder text, List<LegacyDocWritableRun> runs, TableCell cell) {
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
            internal static readonly LegacyDocWritableTableRowFormatting Empty = new LegacyDocWritableTableRowFormatting(null, false);

            internal LegacyDocWritableTableRowFormatting(int? rowHeightTwips, bool rowHeightIsExact) {
                RowHeightTwips = rowHeightTwips;
                RowHeightIsExact = rowHeightIsExact;
            }

            internal int? RowHeightTwips { get; }

            internal bool RowHeightIsExact { get; }
        }
    }
}
