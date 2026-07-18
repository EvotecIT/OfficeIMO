using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTable {

        /// <summary>
        ///     Retrieves a row wrapper at the specified index.
        /// </summary>
        public PowerPointTableRow GetRow(int rowIndex) {
            if (rowIndex < 0 || rowIndex >= Rows) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex));
            }

            A.TableRow tableRow = TableElement.Elements<A.TableRow>().ElementAt(rowIndex);
            return new PowerPointTableRow(this, tableRow);
        }

        /// <summary>
        ///     Retrieves a column wrapper at the specified index.
        /// </summary>
        public PowerPointTableColumn GetColumn(int columnIndex) {
            if (columnIndex < 0 || columnIndex >= Columns) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }

            A.GridColumn column = TableElement.TableGrid!.Elements<A.GridColumn>().ElementAt(columnIndex);
            return new PowerPointTableColumn(this, column);
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
        ///     Adds a new row cloned from a template row.
        /// </summary>
        public PowerPointTableRow AddRowFromTemplate(int templateRowIndex, int? index = null, bool clearText = true) {
            if (templateRowIndex < 0 || templateRowIndex >= Rows) {
                throw new ArgumentOutOfRangeException(nameof(templateRowIndex));
            }

            A.Table table = TableElement;
            A.TableRow templateRow = table.Elements<A.TableRow>().ElementAt(templateRowIndex);
            A.TableRow newRow = (A.TableRow)templateRow.CloneNode(true);
            if (clearText) {
                foreach (A.TableCell cell in newRow.Elements<A.TableCell>()) {
                    ClearCellText(cell);
                }
            }

            int insertAt = index.HasValue ? Math.Min(index.Value, Rows) : Rows;
            if (insertAt < Rows) {
                A.TableRow refRow = table.Elements<A.TableRow>().ElementAt(insertAt);
                table.InsertBefore(newRow, refRow);
            } else {
                table.Append(newRow);
            }

            return new PowerPointTableRow(this, newRow);
        }

        /// <summary>
        ///     Removes a row at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the row to remove.</param>
        public void RemoveRow(int index) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableRow row = table.Elements<A.TableRow>().ElementAt(index);
            string[] discardedSoundIds = PowerPointEmbeddedSound
                .GetRelationshipIds(row);
            row.Remove();
            PowerPointEmbeddedSound.RemoveIfUnused(_slidePart,
                discardedSoundIds);
        }

        /// <summary>
        ///     Adds a new column to the table.
        /// </summary>
        /// <param name="index">Optional zero-based index where the column should be inserted. If omitted, column is appended.</param>
        public void AddColumn(int? index = null) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableGrid grid = table.TableGrid!;
            int existingColumns = Columns;
            A.GridColumn gridColumn = new() { Width = 3708400L };

            if (index.HasValue && index.Value < existingColumns) {
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

                if (index.HasValue && index.Value < existingColumns) {
                    A.TableCell refCell = row.Elements<A.TableCell>().ElementAt(index.Value);
                    row.InsertBefore(cell, refCell);
                } else {
                    row.Append(cell);
                }
            }
        }

        /// <summary>
        ///     Adds a new column cloned from a template column.
        /// </summary>
        public PowerPointTableColumn AddColumnFromTemplate(int templateColumnIndex, int? index = null, bool clearText = true) {
            if (templateColumnIndex < 0 || templateColumnIndex >= Columns) {
                throw new ArgumentOutOfRangeException(nameof(templateColumnIndex));
            }

            A.Table table = TableElement;
            A.TableGrid grid = table.TableGrid ?? throw new InvalidOperationException("Table grid is missing.");
            int existingColumns = Columns;
            A.GridColumn templateColumn = grid.Elements<A.GridColumn>().ElementAt(templateColumnIndex);
            A.GridColumn newColumn = (A.GridColumn)templateColumn.CloneNode(true);

            int insertAt = index.HasValue ? Math.Min(index.Value, existingColumns) : existingColumns;
            if (insertAt < existingColumns) {
                A.GridColumn refColumn = grid.Elements<A.GridColumn>().ElementAt(insertAt);
                grid.InsertBefore(newColumn, refColumn);
            } else {
                grid.Append(newColumn);
            }

            foreach (A.TableRow row in table.Elements<A.TableRow>()) {
                A.TableCell templateCell = row.Elements<A.TableCell>().ElementAt(templateColumnIndex);
                A.TableCell newCell = (A.TableCell)templateCell.CloneNode(true);
                if (clearText) {
                    ClearCellText(newCell);
                }

                if (insertAt < existingColumns) {
                    A.TableCell refCell = row.Elements<A.TableCell>().ElementAt(insertAt);
                    row.InsertBefore(newCell, refCell);
                } else {
                    row.Append(newCell);
                }
            }

            return new PowerPointTableColumn(this, newColumn);
        }

        /// <summary>
        ///     Removes a column at the specified index.
        /// </summary>
        /// <param name="index">Zero-based index of the column to remove.</param>
        public void RemoveColumn(int index) {
            A.Table table = Frame.Graphic!.GraphicData!.GetFirstChild<A.Table>()!;
            A.TableGrid grid = table.TableGrid!;
            A.TableCell[] discardedCells = table.Elements<A.TableRow>()
                .Select(row => row.Elements<A.TableCell>().ElementAt(index))
                .ToArray();
            string[] discardedSoundIds = discardedCells
                .SelectMany(PowerPointEmbeddedSound.GetRelationshipIds)
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            grid.Elements<A.GridColumn>().ElementAt(index).Remove();
            foreach (A.TableCell cell in discardedCells) {
                cell.Remove();
            }
            PowerPointEmbeddedSound.RemoveIfUnused(_slidePart,
                discardedSoundIds);
        }
    }
}
