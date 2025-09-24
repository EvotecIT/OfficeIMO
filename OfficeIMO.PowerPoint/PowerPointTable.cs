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
            A.TableRow tableRow = table.Elements<A.TableRow>().ElementAt(row);
            A.TableCell cell = tableRow.Elements<A.TableCell>().ElementAt(column);
            return new PowerPointTableCell(cell);
        }

        /// <summary>
        ///     Adds a new row to the table.
        /// </summary>
        /// <param name="index">Optional zero-based index where the row should be inserted. If omitted, row is appended.</param>
        public void AddRow(int? index = null) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            int rowCount = Rows;
            int insertIndex = index ?? rowCount;

            if (insertIndex < 0 || insertIndex > rowCount) {
                throw new ArgumentOutOfRangeException(nameof(index), insertIndex,
                    $"Row index must be between 0 and {rowCount} (inclusive). Actual value was {insertIndex}.");
            }

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

            if (insertIndex < rowCount) {
                A.TableRow refRow = table.Elements<A.TableRow>().ElementAt(insertIndex);
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
            int rowCount = Rows;

            if (index < 0 || index >= rowCount) {
                throw new ArgumentOutOfRangeException(nameof(index), index,
                    $"Row index must be between 0 and {rowCount - 1} (inclusive). Actual value was {index}.");
            }

            A.TableRow row = table.Elements<A.TableRow>().ElementAt(index);
            row.Remove();
        }

        /// <summary>
        ///     Adds a new column to the table.
        /// </summary>
        /// <param name="index">Optional zero-based index where the column should be inserted. If omitted, column is appended.</param>
        public void AddColumn(int? index = null) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableGrid grid = table.TableGrid!;
            int columnCount = Columns;
            int insertIndex = index ?? columnCount;

            if (insertIndex < 0 || insertIndex > columnCount) {
                throw new ArgumentOutOfRangeException(nameof(index), insertIndex,
                    $"Column index must be between 0 and {columnCount} (inclusive). Actual value was {insertIndex}.");
            }

            A.GridColumn gridColumn = new() { Width = 3708400L };

            if (insertIndex < columnCount) {
                A.GridColumn refCol = grid.Elements<A.GridColumn>().ElementAt(insertIndex);
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

                if (insertIndex < columnCount) {
                    A.TableCell refCell = row.Elements<A.TableCell>().ElementAt(insertIndex);
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
            int columnCount = Columns;

            if (index < 0 || index >= columnCount) {
                throw new ArgumentOutOfRangeException(nameof(index), index,
                    $"Column index must be between 0 and {columnCount - 1} (inclusive). Actual value was {index}.");
            }

            A.TableGrid grid = table.TableGrid!;
            grid.Elements<A.GridColumn>().ElementAt(index).Remove();
            foreach (A.TableRow row in table.Elements<A.TableRow>()) {
                row.Elements<A.TableCell>().ElementAt(index).Remove();
            }
        }
    }
}