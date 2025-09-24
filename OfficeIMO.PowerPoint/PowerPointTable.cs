using System;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a table on a slide.
    /// </summary>
    public class PowerPointTable : PowerPointShape {
        internal PowerPointTable(GraphicFrame frame, Action onChanged) : base(frame, onChanged) {
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
            return new PowerPointTableCell(cell, ChangeHandler);
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
            NotifyChanged();
        }

        /// <summary>
        ///     Removes a row at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the row to remove.</param>
        public void RemoveRow(int index) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableRow row = table.Elements<A.TableRow>().ElementAt(index);
            row.Remove();
            NotifyChanged();
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
            NotifyChanged();
        }

        /// <summary>
        ///     Removes a column at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the column to remove.</param>
        public void RemoveColumn(int index) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableGrid grid = table.TableGrid!;
            grid.Elements<A.GridColumn>().ElementAt(index).Remove();
            foreach (A.TableRow row in table.Elements<A.TableRow>()) {
                row.Elements<A.TableCell>().ElementAt(index).Remove();
            }
            NotifyChanged();
        }
    }
}