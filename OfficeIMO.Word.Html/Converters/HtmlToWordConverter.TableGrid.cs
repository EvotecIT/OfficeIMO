using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private int DetermineTableColumnCount(IHtmlTableElement tableElem, int rows, HtmlToWordOptions options) {
            var occupied = new HashSet<long>();
            int columns = 0;
            int rowIndex = 0;

            void HandleRows(IHtmlCollection<IHtmlTableRowElement> htmlRows) {
                int groupRowCount = htmlRows.Length;
                for (int localRowIndex = 0; localRowIndex < groupRowCount; localRowIndex++) {
                    IHtmlTableRowElement htmlRow = htmlRows[localRowIndex];
                    int columnIndex = 0;
                    for (int cellIndex = 0; cellIndex < htmlRow.Cells.Length; cellIndex++) {
                        while (occupied.Contains(GetTableGridKey(rowIndex, columnIndex))) {
                            columnIndex++;
                        }

                        var htmlCell = htmlRow.Cells[cellIndex] as IHtmlTableCellElement;
                        int rowSpan = GetHtmlRowSpan(htmlCell, options.MaxTableCells.HasValue);
                        int columnSpan = GetHtmlColumnSpan(htmlCell, options.MaxTableCells.HasValue);
                        if (rowSpan == 0) {
                            rowSpan = groupRowCount - localRowIndex;
                        }

                        rowSpan = Math.Max(1, Math.Min(rowSpan, rows - rowIndex));
                        long proposedColumns = (long)columnIndex + columnSpan;
                        ValidateTableLimit(options, rows, proposedColumns);
                        if (proposedColumns > int.MaxValue) {
                            ThrowLimitExceeded(
                                options,
                                "TableSizeLimitExceeded",
                                "HTML table column count exceeded the native Word table limit.",
                                "WordTableColumns",
                                proposedColumns,
                                int.MaxValue);
                        }

                        int boundedColumns = (int)proposedColumns;
                        columns = Math.Max(columns, boundedColumns);
                        for (int reservedRow = rowIndex; reservedRow < rowIndex + rowSpan && reservedRow < rows; reservedRow++) {
                            for (int reservedColumn = columnIndex; reservedColumn < boundedColumns; reservedColumn++) {
                                occupied.Add(GetTableGridKey(reservedRow, reservedColumn));
                            }
                        }

                        columnIndex = boundedColumns;
                    }

                    rowIndex++;
                }
            }

            if (tableElem.Head != null) {
                HandleRows(tableElem.Head.Rows);
            }

            foreach (IHtmlTableSectionElement body in tableElem.Bodies) {
                HandleRows(body.Rows);
            }

            if (tableElem.Foot != null) {
                HandleRows(tableElem.Foot.Rows);
            }

            return columns;
        }

        private static long GetTableGridKey(int row, int column) {
            return ((long)row << 32) | (uint)column;
        }

        private static int GetHtmlRowSpan(IHtmlTableCellElement? htmlCell, bool useRawAttribute) {
            if (htmlCell == null) {
                return 1;
            }

            if (useRawAttribute &&
                HtmlIntegerSemantics.TryParseNonNegativeInteger(htmlCell.GetAttribute("rowspan"), out int rowSpan)) {
                return rowSpan;
            }

            return htmlCell.RowSpan;
        }

        private static int GetHtmlColumnSpan(IHtmlTableCellElement? htmlCell, bool useRawAttribute) {
            if (htmlCell == null) {
                return 1;
            }

            if (useRawAttribute &&
                HtmlIntegerSemantics.TryParsePositiveInteger(htmlCell.GetAttribute("colspan"), out int columnSpan)) {
                return columnSpan;
            }

            return Math.Max(1, htmlCell.ColumnSpan);
        }
    }
}
