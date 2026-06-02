using System;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal bool TryInsertOwnedDataTableAsDeferredDirectSave(DataTable table, int startRow, bool includeHeaders, string range) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (string.IsNullOrEmpty(range)) {
                return false;
            }

            if (!CanRegisterDirectTabularSaveCandidate(startRow, 1, table.Columns.Count)) {
                return false;
            }

            return _excelDocument.RegisterDeferredDirectTabularSaveCandidate(this, table, includeHeaders, range);
        }

        internal bool TryInsertRowsAsDeferredDirectSave(
            string tableNameForModel,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<Type> columnTypes,
            object?[][] rows,
            int startRow,
            bool includeHeaders,
            string range) {
            if (columnNames == null) throw new ArgumentNullException(nameof(columnNames));
            if (columnTypes == null) throw new ArgumentNullException(nameof(columnTypes));
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            if (string.IsNullOrEmpty(range)) {
                return false;
            }

            if (!CanRegisterDirectTabularSaveCandidate(startRow, 1, columnNames.Count)) {
                return false;
            }

            if (HasDuplicateObjectExportHeaders(columnNames)) {
                return false;
            }

            return _excelDocument.RegisterDeferredDirectTabularSaveCandidate(
                this,
                tableNameForModel,
                columnNames,
                columnTypes,
                rows,
                includeHeaders,
                range);
        }

        private bool TryInsertCellValuesAsDeferredDirectSave(
            string tableNameForModel,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<Type> columnTypes,
            object?[] values,
            int columnCount,
            int rowCount,
            int startRow,
            bool includeHeaders,
            string range) {
            if (columnNames == null) throw new ArgumentNullException(nameof(columnNames));
            if (columnTypes == null) throw new ArgumentNullException(nameof(columnTypes));
            if (values == null) throw new ArgumentNullException(nameof(values));
            if (string.IsNullOrEmpty(range)) {
                return false;
            }

            if (!CanRegisterDirectTabularSaveCandidate(startRow, 1, columnNames.Count)) {
                return false;
            }

            if (HasDuplicateObjectExportHeaders(columnNames)) {
                return false;
            }

            return _excelDocument.RegisterDeferredDirectCellValuesSaveCandidate(
                this,
                tableNameForModel,
                columnNames,
                columnTypes,
                values,
                columnCount,
                rowCount,
                valuesMatchColumnTypes: false,
                includeHeaders,
                range);
        }

        private static bool HasDuplicateObjectExportHeaders(IEnumerable<string> columnNames) {
            return HasDuplicateObjectExportHeaders(columnNames, StringComparer.OrdinalIgnoreCase);
        }

        private static bool HasDuplicateObjectExportHeaders(IEnumerable<string> columnNames, StringComparer comparer) {
            var seen = new HashSet<string>(comparer);
            foreach (var columnName in columnNames) {
                if (!seen.Add(columnName ?? string.Empty)) {
                    return true;
                }
            }

            return false;
        }

        private bool CanRegisterDirectTabularSaveCandidate(int startRow, int startColumn, int columnCount) {
            if (startRow != 1 || startColumn != 1 || columnCount <= 0 || _excelDocument.HasPackagePropertiesDirty) {
                return false;
            }

            var sheets = WorkbookRoot.Sheets;
            if (sheets == null) {
                return false;
            }

            using var sheetEnumerator = sheets.Elements<Sheet>().GetEnumerator();
            if (!sheetEnumerator.MoveNext()
                || !ReferenceEquals(sheetEnumerator.Current, SheetElement)
                || sheetEnumerator.MoveNext()) {
                return false;
            }

            if (SheetElement.State != null && SheetElement.State.Value != SheetStateValues.Visible) {
                return false;
            }

            if (WorksheetPart.DrawingsPart != null || WorksheetPart.WorksheetCommentsPart != null || WorksheetPart.ExternalRelationships.Any()) {
                return false;
            }

            if (WorksheetPart.TableDefinitionParts.Any()) {
                return false;
            }

            var worksheet = WorksheetRoot;
            foreach (var child in worksheet.ChildElements) {
                if (child is not SheetData sheetData) {
                    return false;
                }

                if (sheetData.Elements<Row>().Any(row => row.Elements<Cell>().Any())) {
                    return false;
                }
            }

            return true;
        }

        private static string BuildObjectExportRange(int startRow, int columnCount, int dataRowCount, bool includeHeaders) {
            int rowCount = dataRowCount + (includeHeaders ? 1 : 0);
            if (columnCount <= 0 || rowCount <= 0) {
                return string.Empty;
            }

            return A1.CellReference(startRow, 1) + ":" + A1.CellReference(startRow + rowCount - 1, columnCount);
        }


        private static object?[][] CreateObjectExportRows(IReadOnlyList<string> headers, IReadOnlyList<Dictionary<string, object?>> rows, out Type[] columnTypes) {
            var values = new object?[rows.Count][];
            var inferredColumnTypes = new Type?[headers.Count];
            for (int r = 0; r < rows.Count; r++) {
                var rowValues = new object?[headers.Count];
                for (int c = 0; c < headers.Count; c++) {
                    object? value = rows[r].TryGetValue(headers[c], out var entry) ? entry : null;
                    rowValues[c] = value;
                    UpdateObjectExportColumnType(inferredColumnTypes, c, value);
                }

                values[r] = rowValues;
            }

            columnTypes = CompleteObjectExportColumnTypes(inferredColumnTypes);
            return values;
        }

        private static void UpdateObjectExportColumnType(Type?[] inferredColumnTypes, int columnIndex, object? value) {
            if (value == null || value == DBNull.Value || inferredColumnTypes[columnIndex] == typeof(object)) {
                return;
            }

            Type valueType = value.GetType();
            UpdateObjectExportColumnType(inferredColumnTypes, columnIndex, valueType);
        }

        private static void UpdateObjectExportColumnType(Type?[] inferredColumnTypes, int columnIndex, Type? valueType) {
            if (valueType == null || inferredColumnTypes[columnIndex] == typeof(object)) {
                return;
            }

            Type? inferred = inferredColumnTypes[columnIndex];
            inferredColumnTypes[columnIndex] = inferred == null || inferred == valueType
                ? valueType
                : typeof(object);
        }

        private static Type[] CompleteObjectExportColumnTypes(Type?[] inferredColumnTypes) {
            var columnTypes = new Type[inferredColumnTypes.Length];
            for (int i = 0; i < columnTypes.Length; i++) {
                columnTypes[i] = inferredColumnTypes[i] ?? typeof(object);
            }

            return columnTypes;
        }
    }
}
