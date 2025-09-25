using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a table on a slide.
    /// </summary>
    public class PowerPointTable : PowerPointShape {
        internal PowerPointTable(GraphicFrame frame) : base(frame) {
        }

        private GraphicFrame Frame => (GraphicFrame)Element;

        /// <summary>
        ///     Returns number of rows in the table.
        /// </summary>
        public int Rows => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!.Elements<A.TableRow>().Count();

        /// <summary>
        ///     Returns number of columns in the table.
        /// </summary>
        public int Columns => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!.TableGrid!.Elements<A.GridColumn>()
            .Count();

        /// <summary>
        ///     Retrieves a cell at the specified row and column index.
        /// </summary>
        /// <param name="row">Zero-based row index.</param>
        /// <param name="column">Zero-based column index.</param>
        public PowerPointTableCell GetCell(int row, int column) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            List<A.TableRow> tableRows = table.Elements<A.TableRow>().ToList();
            int rowCount = tableRows.Count;

            if (row < 0 || row >= rowCount) {
                throw new ArgumentOutOfRangeException(nameof(row),
                    $"Row index {row} is out of range. The table contains {rowCount} row{(rowCount == 1 ? string.Empty : "s")}.");
            }

            A.TableRow tableRow = tableRows[row];
            List<A.TableCell> tableCells = tableRow.Elements<A.TableCell>().ToList();
            int columnCount = tableCells.Count;

            if (column < 0 || column >= columnCount) {
                throw new ArgumentOutOfRangeException(nameof(column),
                    $"Column index {column} is out of range. The table contains {columnCount} column{(columnCount == 1 ? string.Empty : "s")}.");
            }

            A.TableCell cell = tableCells[column];
            return new PowerPointTableCell(cell);
        }

        /// <summary>
        ///     Adds a new row to the table.
        /// </summary>
        /// <param name="index">Optional zero-based index where the row should be inserted. If omitted, row is appended.</param>
        public void AddRow(int? index = null) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            int columns = Columns;
            A.TableRow row = new() { Height = 370840L };
            for (int c = 0; c < columns; c++) {
                A.TableCell cell = new(
                    new A.TextBody(new A.BodyProperties(), new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(string.Empty)))),
                    new A.TableCellProperties()
                );
                row.Append(cell);
            }

            if (index.HasValue && index.Value < Rows) {
                A.TableRow refRow = table.Elements<A.TableRow>().ElementAt(index.Value);
                table.InsertBefore(row, refRow);
            } else {
                table.Append(row);
            }
        }

        /// <summary>
        ///     Removes a row at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the row to remove.</param>
        public void RemoveRow(int index) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            List<A.TableRow> rows = table.Elements<A.TableRow>().ToList();

            if (index < 0 || index >= rows.Count) {
                throw new ArgumentOutOfRangeException(nameof(index),
                    $"Row index {index} is out of range. The table contains {rows.Count} row{(rows.Count == 1 ? string.Empty : "s")}.");
            }

            rows[index].Remove();
        }

        /// <summary>
        ///     Adds a new column to the table.
        /// </summary>
        /// <param name="index">Optional zero-based index where the column should be inserted. If omitted, column is appended.</param>
        public void AddColumn(int? index = null) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableGrid grid = table.TableGrid!;
            A.GridColumn gridColumn = new() { Width = 3708400L };

            if (index.HasValue && index.Value < Columns) {
                A.GridColumn refCol = grid.Elements<A.GridColumn>().ElementAt(index.Value);
                grid.InsertBefore(gridColumn, refCol);
            } else {
                grid.Append(gridColumn);
            }

            foreach (A.TableRow row in table.Elements<A.TableRow>()) {
                A.TableCell cell = new(
                    new A.TextBody(new A.BodyProperties(), new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text(string.Empty)))),
                    new A.TableCellProperties()
                );

                if (index.HasValue && index.Value < Columns) {
                    A.TableCell refCell = row.Elements<A.TableCell>().ElementAt(index.Value);
                    row.InsertBefore(cell, refCell);
                } else {
                    row.Append(cell);
                }
            }
        }

        /// <summary>
        ///     Removes a column at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the column to remove.</param>
        public void RemoveColumn(int index) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableGrid grid = table.TableGrid!;
            List<A.GridColumn> columns = grid.Elements<A.GridColumn>().ToList();

            if (index < 0 || index >= columns.Count) {
                throw new ArgumentOutOfRangeException(nameof(index),
                    $"Column index {index} is out of range. The table contains {columns.Count} column{(columns.Count == 1 ? string.Empty : "s")}.");
            }

            columns[index].Remove();
            foreach (A.TableRow row in table.Elements<A.TableRow>()) {
                List<A.TableCell> cells = row.Elements<A.TableCell>().ToList();
                cells[index].Remove();
            }
        }
    }
}