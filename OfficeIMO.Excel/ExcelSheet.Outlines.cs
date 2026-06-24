using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Sets row outline metadata used by Excel grouping.
        /// </summary>
        /// <param name="rowIndex">One-based row index.</param>
        /// <param name="outlineLevel">Outline level from 0 through 7.</param>
        /// <param name="collapsed">Whether the row is shown as collapsed.</param>
        /// <param name="save">Whether to save the worksheet XML immediately.</param>
        public void SetRowOutline(int rowIndex, byte outlineLevel, bool collapsed = false, bool save = true) {
            if (rowIndex <= 0) {
                throw new ArgumentOutOfRangeException(nameof(rowIndex), "Row index must be greater than 0.");
            }

            ValidateDirectOutlineLevel(outlineLevel);
            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetData sheetData = GetOrCreateSheetData();
                Row row = GetOrCreateRowElement(sheetData, rowIndex);
                row.OutlineLevel = outlineLevel;
                row.Collapsed = collapsed ? true : (bool?)null;
                if (save) {
                    WorksheetRoot.Save();
                }
            });
        }

        /// <summary>
        /// Sets column outline metadata used by Excel grouping.
        /// </summary>
        /// <param name="columnIndex">One-based column index.</param>
        /// <param name="outlineLevel">Outline level from 0 through 7.</param>
        /// <param name="collapsed">Whether the column is shown as collapsed.</param>
        /// <param name="save">Whether to save the worksheet XML immediately.</param>
        public void SetColumnOutline(int columnIndex, byte outlineLevel, bool collapsed = false, bool save = true) {
            if (columnIndex <= 0) {
                throw new ArgumentOutOfRangeException(nameof(columnIndex), "Column index must be greater than 0.");
            }

            ValidateDirectOutlineLevel(outlineLevel);
            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                Worksheet worksheet = WorksheetRoot;
                Columns? columns = worksheet.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = worksheet.InsertAt(new Columns(), 0);
                }

                Column? column = columns.Elements<Column>()
                    .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);
                if (column != null) {
                    column = SplitColumn(columns, column, (uint)columnIndex);
                }

                if (column == null) {
                    column = new Column { Min = (uint)columnIndex, Max = (uint)columnIndex };
                    columns.Append(column);
                }

                column.OutlineLevel = outlineLevel;
                column.Collapsed = collapsed ? true : (bool?)null;
                ReorderColumns(columns);
                if (save) {
                    worksheet.Save();
                }
            });
        }

        private static void ValidateDirectOutlineLevel(byte outlineLevel) {
            if (outlineLevel > 7) {
                throw new ArgumentOutOfRangeException(nameof(outlineLevel), "Outline level must be from 0 through 7.");
            }
        }
    }
}
