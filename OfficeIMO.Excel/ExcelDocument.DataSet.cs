using System.Data;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Imports a <see cref="DataSet"/> into the workbook with one worksheet per source <see cref="DataTable"/>.
        /// </summary>
        /// <param name="dataSet">Source dataset.</param>
        /// <param name="createTables">Create an Excel table over each imported DataTable.</param>
        /// <param name="tableStyle">Excel table style to use when <paramref name="createTables"/> is true.</param>
        /// <param name="includeHeaders">Write source column names as the first row.</param>
        /// <param name="includeAutoFilter">Include table AutoFilter dropdowns when creating tables.</param>
        /// <param name="autoFit">Auto-fit imported columns after each sheet is created.</param>
        /// <param name="mode">Optional execution mode override for DataTable writes.</param>
        /// <param name="ct">Cancellation token.</param>
        /// <returns>Import results describing the created worksheets and ranges.</returns>
        public IReadOnlyList<ExcelDataSetImportResult> InsertDataSet(
            DataSet dataSet,
            bool createTables = true,
            TableStyle tableStyle = TableStyle.TableStyleMedium2,
            bool includeHeaders = true,
            bool includeAutoFilter = true,
            bool autoFit = false,
            ExecutionMode? mode = null,
            CancellationToken ct = default) {
            if (dataSet == null) throw new ArgumentNullException(nameof(dataSet));

            var results = new List<ExcelDataSetImportResult>();
            int tableIndex = 1;
            foreach (DataTable table in dataSet.Tables) {
                ct.ThrowIfCancellationRequested();

                string requestedSheetName = string.IsNullOrWhiteSpace(table.TableName)
                    ? "Table" + tableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : table.TableName;
                ExcelSheet sheet = AddWorkSheet(requestedSheetName, SheetNameValidationMode.Sanitize);
                string? requestedTableName = createTables ? requestedSheetName : null;
                string? actualTableName = null;
                string range;
                if (createTables) {
                    range = sheet.InsertDataTableAsTable(
                        table,
                        includeHeaders: includeHeaders,
                        tableName: requestedTableName,
                        style: tableStyle,
                        includeAutoFilter: includeAutoFilter,
                        mode: mode,
                        ct: ct);
                    actualTableName = GetImportedTableName(sheet);
                } else {
                    sheet.InsertDataTable(table, includeHeaders: includeHeaders, mode: mode, ct: ct);
                    range = BuildImportedRange(table.Rows.Count, table.Columns.Count, includeHeaders);
                }

                if (autoFit && table.Columns.Count > 0) {
                    sheet.AutoFitColumnsFor(Enumerable.Range(1, table.Columns.Count));
                }

                results.Add(new ExcelDataSetImportResult(sheet.Name, actualTableName, range, table.Rows.Count, table.Columns.Count));
                tableIndex++;
            }

            return results;
        }

        private static string? GetImportedTableName(ExcelSheet sheet) {
            return sheet.WorksheetPart.TableDefinitionParts
                .Select(part => part.Table?.Name?.Value ?? part.Table?.DisplayName?.Value)
                .FirstOrDefault(name => !string.IsNullOrWhiteSpace(name));
        }

        private static string BuildImportedRange(int rowCount, int columnCount, bool includeHeaders) {
            int rows = rowCount + (includeHeaders ? 1 : 0);
            int columns = Math.Max(1, columnCount);
            if (rows < 1) {
                rows = 1;
            }

            return A1.CellReference(1, 1) + ":" + A1.CellReference(rows, columns);
        }
    }
}
