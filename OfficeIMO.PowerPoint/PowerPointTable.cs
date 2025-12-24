using System;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a table on a slide.
    /// </summary>
    public class PowerPointTable : PowerPointShape {
        private const int EmusPerPoint = 12700;
        internal PowerPointTable(GraphicFrame frame) : base(frame) {
        }

        private GraphicFrame Frame => (GraphicFrame)Element;
        internal A.Table TableElement => Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;

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
        ///     Enables or disables header row styling (firstRow attribute) on the table.
        /// </summary>
        public bool HeaderRow {
            get => TableElement.TableProperties?.FirstRow?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.FirstRow = value;
            }
        }

        /// <summary>
        ///     Enables or disables banded rows styling (bandRow attribute) on the table.
        /// </summary>
        public bool BandedRows {
            get => TableElement.TableProperties?.BandRow?.Value == true;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.BandRow = value;
            }
        }

        /// <summary>
        ///     Gets or sets the table style ID GUID.
        /// </summary>
        public string? StyleId {
            get => TableElement.TableProperties?.GetFirstChild<A.TableStyleId>()?.Text;
            set {
                TableElement.TableProperties ??= new A.TableProperties();
                TableElement.TableProperties.RemoveAllChildren<A.TableStyleId>();
                if (!string.IsNullOrWhiteSpace(value)) {
                    TableElement.TableProperties.Append(new A.TableStyleId { Text = value! });
                }
            }
        }

        /// <summary>
        ///     Sets the width of a specific column in EMUs.
        /// </summary>
        public void SetColumnWidth(int columnIndex, long width) {
            if (columnIndex < 0 || columnIndex >= Columns) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }
            A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(columnIndex);
            column.Width = width;
        }

        /// <summary>
        ///     Sets the width of a specific column in points.
        /// </summary>
        public void SetColumnWidthPoints(int columnIndex, double widthPoints) {
            SetColumnWidth(columnIndex, ToEmus(widthPoints));
        }

        /// <summary>
        ///     Sets widths for columns in points (applies to the first N columns provided).
        /// </summary>
        public void SetColumnWidthsPoints(params double[] widthsPoints) {
            if (widthsPoints == null) {
                throw new ArgumentNullException(nameof(widthsPoints));
            }
            int count = Math.Min(widthsPoints.Length, Columns);
            for (int i = 0; i < count; i++) {
                SetColumnWidthPoints(i, widthsPoints[i]);
            }
        }

        /// <summary>
        ///     Sets the height of a specific row in EMUs.
        /// </summary>
        public void SetRowHeight(int rowIndex, long height) {
            if (rowIndex < 0 || rowIndex >= Rows) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex));
            }
            A.TableRow row = TableElement.Elements<A.TableRow>().ElementAt(rowIndex);
            row.Height = height;
        }

        /// <summary>
        ///     Sets the height of a specific row in points.
        /// </summary>
        public void SetRowHeightPoints(int rowIndex, double heightPoints) {
            SetRowHeight(rowIndex, ToEmus(heightPoints));
        }

        /// <summary>
        ///     Sets heights for rows in points (applies to the first N rows provided).
        /// </summary>
        public void SetRowHeightsPoints(params double[] heightsPoints) {
            if (heightsPoints == null) {
                throw new ArgumentNullException(nameof(heightsPoints));
            }
            int count = Math.Min(heightsPoints.Length, Rows);
            for (int i = 0; i < count; i++) {
                SetRowHeightPoints(i, heightsPoints[i]);
            }
        }

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
            grid.Elements<A.GridColumn>().ElementAt(index).Remove();
            foreach (A.TableRow row in table.Elements<A.TableRow>()) {
                row.Elements<A.TableCell>().ElementAt(index).Remove();
            }
        }

        private static long ToEmus(double points) {
            return (long)Math.Round(points * EmusPerPoint);
        }
    }
}
