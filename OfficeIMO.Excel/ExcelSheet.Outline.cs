using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Groups a contiguous range of worksheet rows by applying an Excel outline level.
        /// </summary>
        /// <param name="startRow">First 1-based row in the group.</param>
        /// <param name="endRow">Last 1-based row in the group.</param>
        /// <param name="outlineLevel">Excel outline level from 1 through 7.</param>
        /// <param name="collapsed">Whether to hide the grouped rows and mark the summary row as collapsed.</param>
        /// <param name="hidden">Whether to hide the grouped rows without marking the group as collapsed.</param>
        public void GroupRows(int startRow, int endRow, byte outlineLevel = 1, bool collapsed = false, bool hidden = false) {
            ValidateOutlineRange(startRow, endRow, nameof(startRow), nameof(endRow));
            ValidateOutlineLevel(outlineLevel);

            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetData sheetData = GetOrCreateSheetData();
                for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
                    Row row = GetOrCreateRowElement(sheetData, rowIndex);
                    row.OutlineLevel = outlineLevel;
                    row.Hidden = collapsed || hidden ? true : null;
                    row.Collapsed = null;
                }

                if (collapsed) {
                    Row summaryRow = GetOrCreateRowElement(sheetData, endRow + 1);
                    summaryRow.Collapsed = true;
                }

                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Clears Excel outline grouping metadata from a contiguous range of worksheet rows.
        /// </summary>
        /// <param name="startRow">First 1-based row to clear.</param>
        /// <param name="endRow">Last 1-based row to clear.</param>
        /// <param name="unhide">Whether hidden rows in the cleared range should be shown.</param>
        public void ClearRowGroup(int startRow, int endRow, bool unhide = true) {
            ValidateOutlineRange(startRow, endRow, nameof(startRow), nameof(endRow));

            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetData? sheetData = WorksheetRoot.GetFirstChild<SheetData>();
                if (sheetData == null) {
                    return;
                }

                for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
                    Row? row = sheetData.Elements<Row>()
                        .FirstOrDefault(candidate => candidate.RowIndex?.Value == (uint)rowIndex);
                    if (row == null) {
                        continue;
                    }

                    row.OutlineLevel = null;
                    row.Collapsed = null;
                    if (unhide) {
                        row.Hidden = null;
                    }

                    RemoveRowWhenEmpty(row);
                }

                Row? nextRow = sheetData.Elements<Row>()
                    .FirstOrDefault(candidate => candidate.RowIndex?.Value == (uint)(endRow + 1));
                if (nextRow != null && nextRow.Collapsed?.Value == true) {
                    nextRow.Collapsed = null;
                    RemoveRowWhenEmpty(nextRow);
                }

                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Groups a contiguous range of worksheet columns by applying an Excel outline level.
        /// </summary>
        /// <param name="startColumn">First 1-based column in the group.</param>
        /// <param name="endColumn">Last 1-based column in the group.</param>
        /// <param name="outlineLevel">Excel outline level from 1 through 7.</param>
        /// <param name="collapsed">Whether to hide the grouped columns and mark the summary column as collapsed.</param>
        /// <param name="hidden">Whether to hide the grouped columns without marking the group as collapsed.</param>
        public void GroupColumns(int startColumn, int endColumn, byte outlineLevel = 1, bool collapsed = false, bool hidden = false) {
            ValidateOutlineRange(startColumn, endColumn, nameof(startColumn), nameof(endColumn));
            ValidateOutlineLevel(outlineLevel);
            if (endColumn >= A1.MaxColumns && collapsed) {
                throw new ArgumentOutOfRangeException(nameof(endColumn), "Collapsed column groups require a following summary column.");
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                Columns columns = GetOrCreateColumns();
                for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++) {
                    Column column = GetOrCreateSingleColumn(columns, columnIndex);
                    column.OutlineLevel = outlineLevel;
                    column.Hidden = collapsed || hidden ? true : null;
                    column.Collapsed = null;
                }

                if (collapsed) {
                    Column summaryColumn = GetOrCreateSingleColumn(columns, endColumn + 1);
                    summaryColumn.Collapsed = true;
                }

                ReorderColumns(columns);
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Clears Excel outline grouping metadata from a contiguous range of worksheet columns.
        /// </summary>
        /// <param name="startColumn">First 1-based column to clear.</param>
        /// <param name="endColumn">Last 1-based column to clear.</param>
        /// <param name="unhide">Whether hidden columns in the cleared range should be shown.</param>
        public void ClearColumnGroup(int startColumn, int endColumn, bool unhide = true) {
            ValidateOutlineRange(startColumn, endColumn, nameof(startColumn), nameof(endColumn));

            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                Columns? columns = WorksheetRoot.GetFirstChild<Columns>();
                if (columns == null) {
                    return;
                }

                for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++) {
                    Column? column = columns.Elements<Column>()
                        .FirstOrDefault(candidate => candidate.Min?.Value <= (uint)columnIndex && candidate.Max?.Value >= (uint)columnIndex);
                    if (column == null) {
                        continue;
                    }

                    column = SplitColumn(columns, column, (uint)columnIndex);
                    column.OutlineLevel = null;
                    column.Collapsed = null;
                    if (unhide) {
                        column.Hidden = null;
                    }

                    RemoveColumnWhenEmpty(column);
                }

                Column? nextColumn = columns.Elements<Column>()
                    .FirstOrDefault(candidate => candidate.Min?.Value <= (uint)(endColumn + 1) && candidate.Max?.Value >= (uint)(endColumn + 1));
                if (nextColumn != null && nextColumn.Collapsed?.Value == true) {
                    nextColumn = SplitColumn(columns, nextColumn, (uint)(endColumn + 1));
                    nextColumn.Collapsed = null;
                    RemoveColumnWhenEmpty(nextColumn);
                }

                if (columns.Elements<Column>().Any()) {
                    ReorderColumns(columns);
                } else {
                    columns.Remove();
                }

                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Configures the worksheet outline summary positions shown by Excel.
        /// </summary>
        /// <param name="summaryBelow">Whether row summaries should appear below grouped rows.</param>
        /// <param name="summaryRight">Whether column summaries should appear to the right of grouped columns.</param>
        public void SetOutlineSummary(bool? summaryBelow = null, bool? summaryRight = null) {
            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetProperties sheetProperties = WorksheetRoot.GetFirstChild<SheetProperties>() ?? new SheetProperties();
                if (sheetProperties.Parent == null) {
                    WorksheetRoot.InsertAt(sheetProperties, 0);
                }

                OutlineProperties outlineProperties = sheetProperties.GetFirstChild<OutlineProperties>() ?? new OutlineProperties();
                if (outlineProperties.Parent == null) {
                    sheetProperties.Append(outlineProperties);
                }

                if (summaryBelow.HasValue) {
                    outlineProperties.SummaryBelow = summaryBelow.Value;
                }

                if (summaryRight.HasValue) {
                    outlineProperties.SummaryRight = summaryRight.Value;
                }

                WorksheetRoot.Save();
            });
        }

        private Columns GetOrCreateColumns() {
            Columns? columns = WorksheetRoot.GetFirstChild<Columns>();
            if (columns != null) {
                return columns;
            }

            SheetData? sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            columns = new Columns();
            if (sheetData != null) {
                WorksheetRoot.InsertBefore(columns, sheetData);
            } else {
                WorksheetRoot.Append(columns);
            }

            return columns;
        }

        private static Column GetOrCreateSingleColumn(Columns columns, int columnIndex) {
            Column? column = columns.Elements<Column>()
                .FirstOrDefault(candidate => candidate.Min?.Value <= (uint)columnIndex && candidate.Max?.Value >= (uint)columnIndex);
            if (column != null) {
                return SplitColumn(columns, column, (uint)columnIndex);
            }

            column = new Column {
                Min = (uint)columnIndex,
                Max = (uint)columnIndex
            };
            columns.Append(column);
            return column;
        }

        private static void ValidateOutlineRange(int start, int end, string startParameter, string endParameter) {
            if (start <= 0) {
                throw new ArgumentOutOfRangeException(startParameter, "Outline range start must be 1 or greater.");
            }

            if (end < start) {
                throw new ArgumentException("Outline range end must be greater than or equal to the start.", endParameter);
            }
        }

        private static void ValidateOutlineLevel(byte outlineLevel) {
            if (outlineLevel < 1 || outlineLevel > 7) {
                throw new ArgumentOutOfRangeException(nameof(outlineLevel), "Excel outline level must be between 1 and 7.");
            }
        }

        private static void RemoveRowWhenEmpty(Row row) {
            if (!row.Elements<Cell>().Any()
                && row.Height == null
                && row.CustomHeight == null
                && row.Hidden == null
                && row.OutlineLevel == null
                && row.Collapsed == null) {
                row.Remove();
            }
        }

        private static void RemoveColumnWhenEmpty(Column column) {
            if (column.Width == null
                && column.CustomWidth == null
                && column.BestFit == null
                && column.Hidden == null
                && column.OutlineLevel == null
                && column.Collapsed == null
                && column.Style == null) {
                column.Remove();
            }
        }
    }
}
