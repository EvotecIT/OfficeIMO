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
        private static readonly ConcurrentDictionary<Type, SimpleObjectExportPlan> SimpleObjectExportPlans = new();

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

            var rows = items as IReadOnlyList<T> ?? items.ToList();
            if (rows.Count == 0) {
                return;
            }

            if (TryInsertSimpleObjectRowsAsDeferredDirectSave(rows, includeHeaders, startRow)) {
                return;
            }

            var list = new List<object?>(rows.Count);
            for (int i = 0; i < rows.Count; i++) {
                list.Add(rows[i]);
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

            string? directSaveRange = null;
            bool hasBlankDisplayHeader = includeHeaders && headers.Any(string.IsNullOrWhiteSpace);
            if (!hasBlankDisplayHeader && CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Count)) {
                try {
                    directSaveRange = BuildObjectExportRange(startRow, headers.Count, flattenedItems.Count, includeHeaders);
                    var columnTypes = InferObjectExportColumnTypes(flattenedItems, headers);
                    var directRows = CreateObjectExportRows(headers, flattenedItems);
                    if (TryInsertRowsAsDeferredDirectSave(Name, headers, columnTypes, directRows, startRow, includeHeaders, directSaveRange)) {
                        return;
                    }
                } catch {
                    // Direct-save registration is opportunistic; fall back to the normal cell path.
                }
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
            CellValues(cells, hasBlankDisplayHeader ? ExecutionMode.Parallel : null);
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

            var rows = items as IReadOnlyList<T> ?? items.ToList();
            if (rows.Count == 0) {
                return;
            }

            var headers = new string[columns.Length];
            var selectors = new Func<T, object?>[columns.Length];
            bool hasBlankDisplayHeader = false;
            for (int c = 0; c < columns.Length; c++) {
                string header = columns[c].Header ?? $"Column{c + 1}";
                headers[c] = header;
                selectors[c] = columns[c].Selector ?? NullObjectSelector;
                if (includeHeaders && string.IsNullOrWhiteSpace(header)) {
                    hasBlankDisplayHeader = true;
                }
            }

            var values = new object?[rows.Count][];
            for (int r = 0; r < rows.Count; r++) {
                var rowValues = new object?[selectors.Length];
                for (int c = 0; c < selectors.Length; c++) {
                    rowValues[c] = selectors[c](rows[r]);
                }

                values[r] = rowValues;
            }

            string? directSaveRange = null;
            if (!hasBlankDisplayHeader && CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Length)) {
                try {
                    if (!HasDuplicateObjectExportHeaders(headers)) {
                        var columnTypes = InferObjectExportColumnTypes(values, headers.Length);
                        directSaveRange = BuildObjectExportRange(startRow, headers.Length, rows.Count, includeHeaders);
                        if (TryInsertRowsAsDeferredDirectSave(Name, headers, columnTypes, values, startRow, includeHeaders, directSaveRange)) {
                            return;
                        }
                    }
                } catch {
                    // Direct-save registration is opportunistic; fall back to the normal cell path.
                }
            }

            int headerRows = includeHeaders ? 1 : 0;
            int totalCellCount = checked((rows.Count + headerRows) * headers.Length);
            var cells = new (int Row, int Column, object Value)[totalCellCount];
            int cellIndex = 0;
            int row = startRow;
            if (includeHeaders) {
                for (int c = 0; c < headers.Length; c++) {
                    cells[cellIndex++] = (row, c + 1, headers[c]);
                }
                row++;
            }

            foreach (var item in rows) {
                for (int c = 0; c < headers.Length; c++) {
                    object value = values[row - startRow - headerRows][c] ?? string.Empty;
                    cells[cellIndex++] = (row, c + 1, value);
                }
                row++;
            }

            CellValues(cells, hasBlankDisplayHeader ? ExecutionMode.Parallel : null);
        }

        private static object? NullObjectSelector<T>(T row) => null;

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

        private static bool HasDuplicateObjectExportHeaders(IEnumerable<string> columnNames) {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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

        private bool TryInsertSimpleObjectRowsAsDeferredDirectSave<T>(
            IReadOnlyList<T> rows,
            bool includeHeaders,
            int startRow) {
            if (rows.Count == 0) {
                return false;
            }

            Type rowType = rows[0]?.GetType() ?? typeof(object);
            if (rowType == typeof(object)) {
                return false;
            }

            SimpleObjectExportPlan plan = GetSimpleObjectExportPlan(rowType);
            if (!plan.CanUseDirectSave) {
                return false;
            }

            string[] headers = plan.Headers;
            if (!CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Length)) {
                return false;
            }

            SimpleObjectExportValueGetter[] getters = plan.Getters;
            var values = new object?[rows.Count][];
            for (int r = 0; r < rows.Count; r++) {
                object? row = rows[r];
                if (row == null || row.GetType() != rowType) {
                    return false;
                }

                var rowValues = new object?[getters.Length];
                for (int c = 0; c < getters.Length; c++) {
                    rowValues[c] = getters[c](row);
                }

                values[r] = rowValues;
            }

            string range = BuildObjectExportRange(startRow, headers.Length, rows.Count, includeHeaders);
            return TryInsertRowsAsDeferredDirectSave(Name, headers, plan.ColumnTypes, values, startRow, includeHeaders, range);
        }

        private static SimpleObjectExportPlan GetSimpleObjectExportPlan(Type type)
            => SimpleObjectExportPlans.GetOrAdd(type, CreateSimpleObjectExportPlan);

        private static SimpleObjectExportPlan CreateSimpleObjectExportPlan(Type type) {
            var properties = GetSimpleObjectExportProperties(type);
            if (properties.Length == 0) {
                return SimpleObjectExportPlan.NotSupported;
            }

            var headers = new string[properties.Length];
            var getters = new SimpleObjectExportValueGetter[properties.Length];
            for (int i = 0; i < properties.Length; i++) {
                headers[i] = properties[i].Name;
                getters[i] = CreateSimpleObjectExportValueGetter(properties[i]);
            }

            if (HasDuplicateObjectExportHeaders(headers)) {
                return SimpleObjectExportPlan.NotSupported;
            }

            return new SimpleObjectExportPlan(headers, getters, InferSimpleObjectExportColumnTypes(properties), canUseDirectSave: true);
        }

        private static SimpleObjectExportValueGetter CreateSimpleObjectExportValueGetter(PropertyInfo property) {
            MethodInfo? getMethod = property.GetMethod;
            if (getMethod == null || property.DeclaringType == null) {
                return row => property.GetValue(row, null);
            }

            try {
                return (SimpleObjectExportValueGetter)CreateSimpleObjectExportValueGetterMethod
                    .MakeGenericMethod(property.DeclaringType, property.PropertyType)
                    .Invoke(null, new object[] { getMethod })!;
            } catch {
                return row => property.GetValue(row, null);
            }
        }

        private static readonly MethodInfo CreateSimpleObjectExportValueGetterMethod =
            typeof(ExcelSheet).GetMethod(nameof(CreateSimpleObjectExportValueGetterCore), BindingFlags.NonPublic | BindingFlags.Static)!;

        private static SimpleObjectExportValueGetter CreateSimpleObjectExportValueGetterCore<TTarget, TValue>(MethodInfo getMethod) {
            var getter = (Func<TTarget, TValue>)Delegate.CreateDelegate(typeof(Func<TTarget, TValue>), getMethod);
            return row => getter((TTarget)row!);
        }

        private static PropertyInfo[] GetSimpleObjectExportProperties(Type type) {
            var properties = type.GetProperties().Where(property => property.CanRead).ToArray();
            if (properties.Length == 0) {
                return Array.Empty<PropertyInfo>();
            }

            for (int i = 0; i < properties.Length; i++) {
                if (properties[i].GetIndexParameters().Length != 0
                    || !IsSimpleObjectExportScalarType(properties[i].PropertyType)) {
                    return Array.Empty<PropertyInfo>();
                }
            }

            return properties;
        }

        private static Type[] InferSimpleObjectExportColumnTypes(IReadOnlyList<PropertyInfo> properties) {
            var columnTypes = new Type[properties.Count];
            for (int i = 0; i < columnTypes.Length; i++) {
                columnTypes[i] = Nullable.GetUnderlyingType(properties[i].PropertyType) ?? properties[i].PropertyType;
            }

            return columnTypes;
        }

        private static bool IsSimpleObjectExportScalarType(Type type) {
            type = Nullable.GetUnderlyingType(type) ?? type;
            return type.IsPrimitive
                || type == typeof(string)
                || type == typeof(decimal)
                || type == typeof(DateTime)
                || type == typeof(DateTimeOffset)
                || type == typeof(Guid);
        }

        private sealed class SimpleObjectExportPlan {
            internal static readonly SimpleObjectExportPlan NotSupported = new(
                Array.Empty<string>(),
                Array.Empty<SimpleObjectExportValueGetter>(),
                Array.Empty<Type>(),
                canUseDirectSave: false);

            internal SimpleObjectExportPlan(
                string[] headers,
                SimpleObjectExportValueGetter[] getters,
                Type[] columnTypes,
                bool canUseDirectSave) {
                Headers = headers;
                Getters = getters;
                ColumnTypes = columnTypes;
                CanUseDirectSave = canUseDirectSave;
            }

            internal string[] Headers { get; }

            internal SimpleObjectExportValueGetter[] Getters { get; }

            internal Type[] ColumnTypes { get; }

            internal bool CanUseDirectSave { get; }
        }

        private delegate object? SimpleObjectExportValueGetter(object? row);

        private static object?[][] CreateObjectExportRows(IReadOnlyList<string> headers, IReadOnlyList<Dictionary<string, object?>> rows) {
            var values = new object?[rows.Count][];
            for (int r = 0; r < rows.Count; r++) {
                var rowValues = new object?[headers.Count];
                for (int c = 0; c < headers.Count; c++) {
                    rowValues[c] = rows[r].TryGetValue(headers[c], out var entry) ? entry : null;
                }

                values[r] = rowValues;
            }

            return values;
        }

        private static Type[] InferObjectExportColumnTypes(IReadOnlyList<object?[]> values, int columnCount) {
            var columnTypes = new Type[columnCount];
            for (int c = 0; c < columnCount; c++) {
                columnTypes[c] = InferObjectExportColumnType(values, c);
            }

            return columnTypes;
        }

        private static Type[] InferObjectExportColumnTypes(IReadOnlyList<Dictionary<string, object?>> rows, IReadOnlyList<string> headers) {
            var columnTypes = new Type[headers.Count];
            for (int c = 0; c < headers.Count; c++) {
                columnTypes[c] = InferObjectExportColumnType(rows, headers[c]);
            }

            return columnTypes;
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
                cell.CellValue = new CellValue(SharedStringIndexText.Get(sharedStringIndex));
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

