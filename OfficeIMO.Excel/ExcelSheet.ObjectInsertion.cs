using System;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Inserts objects into the worksheet by flattening their properties into columns.
        /// </summary>
        /// <typeparam name="T">Type of objects being inserted.</typeparam>
        /// <param name="items">Collection of objects to insert.</param>
        /// <param name="includeHeaders">Whether to include column headers.</param>
        /// <param name="startRow">1-based starting row.</param>
        [RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly with CellValues or pre-flatten using known types.")]
        public void InsertObjects<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(IEnumerable<T> items, bool includeHeaders = true, int startRow = 1) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            var list = items.Cast<object?>().ToList();
            if (list.Count == 0) {
                return;
            }

            var flattenedItems = new List<Dictionary<string, object?>>(list.Count);
            List<string> headers = new List<string>();
            HashSet<string> headerSet = new HashSet<string>();

            foreach (var item in list) {
                var dict = new Dictionary<string, object?>();
                FlattenObject(item, null, dict);
                flattenedItems.Add(dict);
                foreach (var key in dict.Keys) {
                    if (headerSet.Add(key)) {
                        headers.Add(key);
                    }
                }
            }

            DataTable? directSaveTable = null;
            string? directSaveRange = null;
            bool canRegisterDirectSave = CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Count);
            if (canRegisterDirectSave) {
                try {
                    directSaveTable = CreateObjectExportTable(headers, flattenedItems, Name);
                    directSaveRange = BuildObjectExportRange(startRow, headers.Count, flattenedItems.Count, includeHeaders);
                } catch {
                    canRegisterDirectSave = false;
                }
            }

            if (TryInsertObjectExportTableAndRegisterDirectSave(directSaveTable, includeHeaders, startRow, directSaveRange, canRegisterDirectSave)) {
                return;
            }

            int headerRows = includeHeaders ? 1 : 0;
            int totalCellCount = checked((list.Count + headerRows) * Math.Max(1, headers.Count));
            var cells = new (int Row, int Column, object Value)[totalCellCount];
            int cellIndex = 0;
            int row = startRow;
            if (includeHeaders) {
                for (int c = 0; c < headers.Count; c++) {
                    cells[cellIndex++] = (row, c + 1, headers[c]);
                }
                row++;
            }

            foreach (var dict in flattenedItems) {
                for (int c = 0; c < headers.Count; c++) {
                    object value = dict.TryGetValue(headers[c], out var entry) ? entry ?? string.Empty : string.Empty;
                    cells[cellIndex++] = (row, c + 1, value);
                }
                row++;
            }

            // Use the batch CellValues path with planner + execution policy    
            CellValues(cells, null);
            if (canRegisterDirectSave && directSaveTable != null && !string.IsNullOrEmpty(directSaveRange)) {
                _excelDocument.RegisterDirectTabularSaveCandidate(this, directSaveTable, includeHeaders, directSaveRange!, copyTable: false);
            }
        }

        /// <summary>
        /// Inserts objects into the worksheet using explicit column selectors (AOT-safe).
        /// </summary>
        /// <typeparam name="T">Type of objects being inserted.</typeparam>
        /// <param name="items">Collection of objects to insert.</param>
        /// <param name="columns">Column headers and selectors.</param>
        public void InsertObjects<T>(IEnumerable<T> items, params (string Header, Func<T, object?> Selector)[] columns) {
            InsertObjects(items, includeHeaders: true, startRow: 1, columns);
        }

        /// <summary>
        /// Inserts objects into the worksheet using explicit column selectors (AOT-safe).
        /// </summary>
        /// <typeparam name="T">Type of objects being inserted.</typeparam>
        /// <param name="items">Collection of objects to insert.</param>
        /// <param name="includeHeaders">Whether to include column headers.</param>
        /// <param name="startRow">1-based starting row.</param>
        /// <param name="columns">Column headers and selectors.</param>
        public void InsertObjects<T>(IEnumerable<T> items, bool includeHeaders, int startRow, params (string Header, Func<T, object?> Selector)[] columns) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }
            if (columns == null || columns.Length == 0) {
                throw new ArgumentException("At least one column selector is required.", nameof(columns));
            }

            var rows = items as IList<T> ?? items.ToList();
            if (rows.Count == 0) {
                return;
            }

            var normalizedColumns = new (string Header, Func<T, object?> Selector)[columns.Length];
            for (int c = 0; c < columns.Length; c++) {
                normalizedColumns[c] = (columns[c].Header ?? $"Column{c + 1}", columns[c].Selector ?? (_ => null));
            }

            var values = new object?[rows.Count][];
            for (int r = 0; r < rows.Count; r++) {
                var rowValues = new object?[normalizedColumns.Length];
                for (int c = 0; c < normalizedColumns.Length; c++) {
                    rowValues[c] = normalizedColumns[c].Selector(rows[r]);
                }

                values[r] = rowValues;
            }

            DataTable? directSaveTable = null;
            string? directSaveRange = null;
            bool hasBlankDisplayHeader = includeHeaders && normalizedColumns.Any(column => string.IsNullOrWhiteSpace(column.Header));
            bool canRegisterDirectSave = !hasBlankDisplayHeader && CanRegisterDirectTabularSaveCandidate(startRow, 1, normalizedColumns.Length);
            if (canRegisterDirectSave) {
                try {
                    directSaveTable = CreateObjectExportTable(normalizedColumns.Select(column => column.Header).ToList(), values, Name);
                    directSaveRange = BuildObjectExportRange(startRow, normalizedColumns.Length, rows.Count, includeHeaders);
                } catch {
                    canRegisterDirectSave = false;
                }
            }

            if (TryInsertObjectExportTableAndRegisterDirectSave(directSaveTable, includeHeaders, startRow, directSaveRange, canRegisterDirectSave)) {
                return;
            }

            int headerRows = includeHeaders ? 1 : 0;
            int totalCellCount = checked((rows.Count + headerRows) * normalizedColumns.Length);
            var cells = new (int Row, int Column, object Value)[totalCellCount];
            int cellIndex = 0;
            int row = startRow;
            if (includeHeaders) {
                for (int c = 0; c < normalizedColumns.Length; c++) {
                    cells[cellIndex++] = (row, c + 1, normalizedColumns[c].Header);
                }
                row++;
            }

            foreach (var item in rows) {
                for (int c = 0; c < normalizedColumns.Length; c++) {
                    object value = values[row - startRow - headerRows][c] ?? string.Empty;
                    cells[cellIndex++] = (row, c + 1, value);
                }
                row++;
            }

            CellValues(cells, hasBlankDisplayHeader ? ExecutionMode.Parallel : null);
            if (canRegisterDirectSave && directSaveTable != null && !string.IsNullOrEmpty(directSaveRange)) {
                _excelDocument.RegisterDirectTabularSaveCandidate(this, directSaveTable, includeHeaders, directSaveRange!, copyTable: false);
            }
        }

        private bool TryInsertObjectExportTableAndRegisterDirectSave(DataTable? table, bool includeHeaders, int startRow, string? range, bool canRegisterDirectSave) {
            if (!canRegisterDirectSave || table == null || string.IsNullOrEmpty(range)) {
                return false;
            }

            InsertOwnedDataTable(table, startRow, startColumn: 1, includeHeaders: includeHeaders, registerDirectSaveCandidate: false);
            _excelDocument.RegisterDirectTabularSaveCandidate(this, table, includeHeaders, range!, copyTable: false);
            return true;
        }

        private bool CanRegisterDirectTabularSaveCandidate(int startRow, int startColumn, int columnCount) {
            if (startRow != 1 || startColumn != 1 || columnCount <= 0 || _excelDocument.HasPackagePropertiesDirty) {
                return false;
            }

            var sheets = WorkbookRoot.Sheets?.Elements<Sheet>().ToList();
            if (sheets == null || sheets.Count != 1 || !ReferenceEquals(sheets[0], SheetElement)) {
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

        private static DataTable CreateObjectExportTable(IReadOnlyList<string> headers, IReadOnlyList<Dictionary<string, object?>> rows, string tableName) {
            var table = new DataTable(string.IsNullOrWhiteSpace(tableName) ? "Data" : tableName) {
                Locale = CultureInfo.InvariantCulture
            };

            for (int c = 0; c < headers.Count; c++) {
                string header = string.IsNullOrWhiteSpace(headers[c]) ? "Column" + (c + 1).ToString(CultureInfo.InvariantCulture) : headers[c];
                table.Columns.Add(header, InferObjectExportColumnType(rows, headers[c]));
            }

            table.BeginLoadData();
            try {
                for (int r = 0; r < rows.Count; r++) {
                    var values = new object?[headers.Count];
                    for (int c = 0; c < headers.Count; c++) {
                        object? value = rows[r].TryGetValue(headers[c], out var entry) ? entry : null;
                        values[c] = value ?? DBNull.Value;
                    }

                    table.Rows.Add(values);
                }
            } finally {
                table.EndLoadData();
            }

            return table;
        }

        private static DataTable CreateObjectExportTable(IReadOnlyList<string> headers, IReadOnlyList<object?[]> values, string tableName) {
            var table = new DataTable(string.IsNullOrWhiteSpace(tableName) ? "Data" : tableName) {
                Locale = CultureInfo.InvariantCulture
            };

            for (int c = 0; c < headers.Count; c++) {
                string header = string.IsNullOrWhiteSpace(headers[c]) ? "Column" + (c + 1).ToString(CultureInfo.InvariantCulture) : headers[c];
                table.Columns.Add(header, InferObjectExportColumnType(values, c));
            }

            table.BeginLoadData();
            try {
                for (int r = 0; r < values.Count; r++) {
                    var rowValues = new object?[headers.Count];
                    for (int c = 0; c < headers.Count; c++) {
                        object? value = values[r][c];
                        rowValues[c] = value ?? DBNull.Value;
                    }

                    table.Rows.Add(rowValues);
                }
            } finally {
                table.EndLoadData();
            }

            return table;
        }

        private static Type InferObjectExportColumnType(IReadOnlyList<object?[]> values, int columnIndex) {
            Type? inferred = null;
            for (int r = 0; r < values.Count; r++) {
                object? value = values[r][columnIndex];
                if (value == null || value == DBNull.Value) {
                    continue;
                }

                Type valueType = Nullable.GetUnderlyingType(value.GetType()) ?? value.GetType();
                if (inferred == null) {
                    inferred = valueType;
                    continue;
                }

                if (inferred != valueType) {
                    return typeof(object);
                }
            }

            return inferred ?? typeof(object);
        }

        private static Type InferObjectExportColumnType(IReadOnlyList<Dictionary<string, object?>> rows, string header) {
            Type? inferred = null;
            for (int r = 0; r < rows.Count; r++) {
                object? value = rows[r].TryGetValue(header, out var entry) ? entry : null;
                if (value == null || value == DBNull.Value) {
                    continue;
                }

                Type valueType = Nullable.GetUnderlyingType(value.GetType()) ?? value.GetType();
                if (inferred == null) {
                    inferred = valueType;
                    continue;
                }

                if (inferred != valueType) {
                    return typeof(object);
                }
            }

            return inferred ?? typeof(object);
        }

        private static void FlattenObject(object? value, string? prefix, IDictionary<string, object?> result) {
            if (value == null) {
                if (!string.IsNullOrEmpty(prefix)) {
                    result[prefix!] = null;
                }
                return;
            }

            if (value is IDictionary dictionary) {
                foreach (DictionaryEntry entry in dictionary) {
                    string key = entry.Key?.ToString() ?? string.Empty;
                    string childPrefix = string.IsNullOrEmpty(prefix) ? key : prefix + "." + key;
                    FlattenObject(entry.Value, childPrefix, result);
                }
                return;
            }

            if (value is IEnumerable enumerable && value is not string) {
                var values = new List<string>();
                foreach (var item in enumerable) {
                    values.Add(item?.ToString() ?? string.Empty);
                }
                if (!string.IsNullOrEmpty(prefix)) {
                    result[prefix!] = string.Join(", ", values);
                }
                return;
            }

            Type type = value.GetType();
            if (type.IsPrimitive || value is string || value is decimal || value is DateTime || value is DateTimeOffset || value is Guid) {
                if (!string.IsNullOrEmpty(prefix)) {
                    result[prefix!] = value;
                }
                return;
            }

            var props = type.GetProperties().Where(p => p.CanRead);
            bool hasAny = false;
            foreach (var prop in props) {
                hasAny = true;
                string childPrefix = string.IsNullOrEmpty(prefix) ? prop.Name : prefix + "." + prop.Name;
                FlattenObject(prop.GetValue(value, null), childPrefix, result);
            }

            if (!hasAny && !string.IsNullOrEmpty(prefix)) {
                result[prefix!] = value.ToString();
            }
        }

        private class CellUpdate {
            public int Row { get; }
            public int Column { get; }
            public string Text { get; }
            public DocumentFormat.OpenXml.Spreadsheet.CellValues DataType { get; }
            public bool IsSharedString { get; }

            public CellUpdate(int row, int column, string text, DocumentFormat.OpenXml.Spreadsheet.CellValues dataType, bool isSharedString) {
                Row = row;
                Column = column;
                Text = text;
                DataType = dataType;
                IsSharedString = isSharedString;
            }
        }

        /// <summary>
        /// Obsolete. Use <see cref="CellValues(IEnumerable{ValueTuple{int, int, object}}, ExecutionMode?, System.Threading.CancellationToken)"/>
        /// with <see cref="ExecutionMode.Parallel"/> instead.
        /// </summary>
        /// <param name="cells">Collection of cell coordinates and values.</param>
        [Obsolete("Use CellValues(..., ExecutionMode.Parallel) instead.")]
        public void CellValuesParallel(IEnumerable<(int Row, int Column, object Value)> cells) {
            if (cells == null) {
                throw new ArgumentNullException(nameof(cells));
            }

            CellValues(cells, ExecutionMode.Parallel);
        }

        private CellUpdate PrepareCellUpdate(int row, int column, object value) {
            var dateTimeOffsetStrategy = _excelDocument.DateTimeOffsetWriteStrategy;
            var (cellValue, dataType) = CoerceValueHelper.Coerce(
                value,
                s => new CellValue(s),
                dateTimeOffsetStrategy);

            bool isSharedString = dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
            return new CellUpdate(row, column, cellValue.Text ?? string.Empty, dataType, isSharedString);
        }

        private void ApplyCellUpdate(CellUpdate update) {
            Cell cell = GetCell(update.Row, update.Column);
            if (update.IsSharedString) {
                int sharedStringIndex = _excelDocument.GetSharedStringIndex(update.Text);
                cell.CellValue = new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture));
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
            } else {
                cell.CellValue = new CellValue(update.Text);
                cell.DataType = update.DataType;
            }
        }

        private void ValidateWorksheetXml()
            => WorksheetIntegrityValidator.Validate(_worksheetPart, EffectiveExecution, Name);
    }
}

