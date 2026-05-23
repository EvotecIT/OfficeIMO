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
        private static readonly ConcurrentDictionary<Type, PowerShellObjectExportPlan> PowerShellObjectExportPlans = new();

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

            if (TryInsertSimpleObjectRowsAsCellValues(rows, includeHeaders, startRow)) {
                return;
            }

            if (TryInsertFlatDictionaryRowsAsDeferredDirectSave(rows, includeHeaders, startRow)) {
                return;
            }

            var flattenedItems = new List<Dictionary<string, object?>>(rows.Count);
            List<string> headers = new List<string>();
            HashSet<string> headerSet = new HashSet<string>();

            foreach (var item in rows) {
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
                    var directRows = CreateObjectExportRows(headers, flattenedItems, out var columnTypes);
                    if (TryInsertRowsAsDeferredDirectSave(Name, headers, columnTypes, directRows, startRow, includeHeaders, directSaveRange)) {
                        return;
                    }
                } catch {
                    // Direct-save registration is opportunistic; fall back to the normal cell path.
                }
            }

            int headerRows = includeHeaders ? 1 : 0;
            int totalCellCount = checked((rows.Count + headerRows) * Math.Max(1, headers.Count));
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

            object?[][]? values = null;
            if (!hasBlankDisplayHeader
                && !HasDuplicateObjectExportHeaders(headers)
                && CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Length)) {
                Type?[] inferredColumnTypes;
                values = CreateExplicitObjectExportRows(rows, selectors, out inferredColumnTypes);
                try {
                    var columnTypes = CompleteObjectExportColumnTypes(inferredColumnTypes);
                    string directSaveRange = BuildObjectExportRange(startRow, headers.Length, rows.Count, includeHeaders);
                    if (TryInsertRowsAsDeferredDirectSave(Name, headers, columnTypes, values, startRow, includeHeaders, directSaveRange)) {
                        return;
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

            if (values != null) {
                foreach (object?[] rowValues in values) {
                    for (int c = 0; c < headers.Length; c++) {
                        cells[cellIndex++] = (row, c + 1, rowValues[c] ?? string.Empty);
                    }

                    row++;
                }
            } else {
                foreach (var item in rows) {
                    for (int c = 0; c < headers.Length; c++) {
                        cells[cellIndex++] = (row, c + 1, selectors[c](item) ?? string.Empty);
                    }

                    row++;
                }
            }

            CellValues(cells, hasBlankDisplayHeader ? ExecutionMode.Parallel : null);
        }

        private static object? NullObjectSelector<T>(T row) => null;

        private static object?[][] CreateExplicitObjectExportRows<T>(
            IReadOnlyList<T> rows,
            IReadOnlyList<Func<T, object?>> selectors,
            out Type?[] inferredColumnTypes) {
            var values = new object?[rows.Count][];
            inferredColumnTypes = new Type?[selectors.Count];
            for (int r = 0; r < rows.Count; r++) {
                var rowValues = new object?[selectors.Count];
                for (int c = 0; c < selectors.Count; c++) {
                    object? value = selectors[c](rows[r]);
                    rowValues[c] = value;
                    UpdateObjectExportColumnType(inferredColumnTypes, c, value);
                }

                values[r] = rowValues;
            }

            return values;
        }

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

        private bool TryInsertSimpleObjectRowsAsDeferredDirectSave<T>(
            IReadOnlyList<T> rows,
            bool includeHeaders,
            int startRow) {
            if (rows.Count == 0) {
                return false;
            }

            bool requireRuntimeTypeCheck = !typeof(T).IsValueType && !typeof(T).IsSealed;
            Type rowType = requireRuntimeTypeCheck ? rows[0]?.GetType() ?? typeof(object) : typeof(T);
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
                if (row == null || requireRuntimeTypeCheck && row.GetType() != rowType) {
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

        private bool TryInsertSimpleObjectRowsAsCellValues<T>(
            IReadOnlyList<T> rows,
            bool includeHeaders,
            int startRow) {
            if (rows.Count == 0) {
                return false;
            }

            bool requireRuntimeTypeCheck = !typeof(T).IsValueType && !typeof(T).IsSealed;
            Type rowType = requireRuntimeTypeCheck ? rows[0]?.GetType() ?? typeof(object) : typeof(T);
            if (rowType == typeof(object)) {
                return false;
            }

            SimpleObjectExportPlan plan = GetSimpleObjectExportPlan(rowType);
            if (!plan.CanUseDirectSave) {
                return false;
            }

            string[] headers = plan.Headers;
            SimpleObjectExportValueGetter[] getters = plan.Getters;
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

            for (int r = 0; r < rows.Count; r++) {
                object? item = rows[r];
                if (item == null || requireRuntimeTypeCheck && item.GetType() != rowType) {
                    return false;
                }

                for (int c = 0; c < getters.Length; c++) {
                    cells[cellIndex++] = (row, c + 1, getters[c](item) ?? string.Empty);
                }

                row++;
            }

            CellValues(cells);
            return true;
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
                    || !IsObjectExportScalarType(properties[i].PropertyType)) {
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

        private static bool IsObjectExportScalarType(Type type) {
            type = Nullable.GetUnderlyingType(type) ?? type;
            return type.IsPrimitive
                || type.IsEnum
                || type == typeof(string)
                || type == typeof(decimal)
                || type == typeof(DateTime)
                || type == typeof(DateTimeOffset)
                || type == typeof(TimeSpan)
                || type == typeof(Guid)
#if NET6_0_OR_GREATER
                || type == typeof(DateOnly)
                || type == typeof(TimeOnly)
#endif
                ;
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

        private bool TryInsertFlatDictionaryRowsAsDeferredDirectSave<T>(
            IReadOnlyList<T> rows,
            bool includeHeaders,
            int startRow) {
            if (rows.Count == 0) {
                return false;
            }

            if (!CanRegisterDirectTabularSaveCandidate(startRow, 1, columnCount: 1)) {
                return false;
            }

            if (TryInsertExactDictionaryRowsAsDeferredDirectSave(rows, includeHeaders, startRow)) {
                return true;
            }

            if (TryInsertReadOnlyDictionaryRowsAsDeferredDirectSave(rows, includeHeaders, startRow)) {
                return true;
            }

            if (TryInsertLegacyDictionaryRowsAsDeferredDirectSave(rows, includeHeaders, startRow)) {
                return true;
            }

            var headers = new List<string>();
            var headerIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
            var directRows = new object?[rows.Count][];
            var state = new FlatDictionaryProjectionState();

            for (int r = 0; r < rows.Count; r++) {
                if (!TryProjectFlatDictionaryRow(
                    rows[r],
                    r,
                    headers,
                    headerIndexes,
                    directRows,
                    state)) {
                    return false;
                }
            }

            NormalizeFlatDictionaryRowWidths(directRows, headers.Count);
            state.NormalizeColumnTypeWidth(headers.Count);

            if (headers.Count == 0
                || includeHeaders && state.HasBlankDisplayHeader
                || HasDuplicateObjectExportHeaders(headers)
                || !CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Count)) {
                return false;
            }

            string range = BuildObjectExportRange(startRow, headers.Count, directRows.Length, includeHeaders);
            return TryInsertRowsAsDeferredDirectSave(
                Name,
                headers,
                CompleteObjectExportColumnTypes(state.InferredColumnTypes),
                directRows,
                startRow,
                includeHeaders,
                range);
        }

        private bool TryInsertExactDictionaryRowsAsDeferredDirectSave<T>(
            IReadOnlyList<T> rows,
            bool includeHeaders,
            int startRow) {
            IReadOnlyList<Dictionary<string, object?>> dictionaryRows;
            if (rows is IReadOnlyList<Dictionary<string, object?>> typedDictionaryRows) {
                dictionaryRows = typedDictionaryRows;
            } else {
                var copiedRows = new Dictionary<string, object?>[rows.Count];
                for (int r = 0; r < copiedRows.Length; r++) {
                    if (rows[r] is not Dictionary<string, object?> dictionary) {
                        return false;
                    }

                    copiedRows[r] = dictionary;
                }

                dictionaryRows = copiedRows;
            }

            var headers = new List<string>();
            var headerIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
            var state = new FlatDictionaryProjectionState();

            for (int r = 0; r < dictionaryRows.Count; r++) {
                Dictionary<string, object?> dictionary = dictionaryRows[r];
                foreach (var entry in dictionary) {
                    if (!IsFlatDictionaryObjectExportValue(entry.Value)) {
                        return false;
                    }

                    string columnName = entry.Key ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(columnName)) {
                        state.HasBlankDisplayHeader = true;
                    }

                    if (!headerIndexes.TryGetValue(columnName, out int columnIndex)) {
                        columnIndex = headers.Count;
                        headers.Add(columnName);
                        headerIndexes.Add(columnName, columnIndex);
                        state.EnsureColumnTypeCapacity(headers.Count);
                    }

                    UpdateObjectExportColumnType(state.InferredColumnTypes, columnIndex, entry.Value);
                }
            }

            state.NormalizeColumnTypeWidth(headers.Count);

            if (headers.Count == 0
                || includeHeaders && state.HasBlankDisplayHeader
                || HasDuplicateObjectExportHeaders(headers)
                || !CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Count)) {
                return false;
            }

            string range = BuildObjectExportRange(startRow, headers.Count, rows.Count, includeHeaders);
            return _excelDocument.RegisterDeferredDirectExactDictionaryRowsSaveCandidate(
                this,
                Name,
                headers,
                CompleteObjectExportColumnTypes(state.InferredColumnTypes),
                dictionaryRows,
                includeHeaders,
                range);
        }

        private bool TryInsertReadOnlyDictionaryRowsAsDeferredDirectSave<T>(
            IReadOnlyList<T> rows,
            bool includeHeaders,
            int startRow) {
            var headers = new List<string>();
            var headerIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
            var dictionaryRows = new IReadOnlyDictionary<string, object?>[rows.Count];
            var state = new FlatDictionaryProjectionState();

            for (int r = 0; r < rows.Count; r++) {
                if (rows[r] is not IReadOnlyDictionary<string, object?> dictionary) {
                    return false;
                }

                dictionaryRows[r] = dictionary;
                foreach (var entry in dictionary) {
                    if (!IsFlatDictionaryObjectExportValue(entry.Value)) {
                        return false;
                    }

                    string columnName = entry.Key ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(columnName)) {
                        state.HasBlankDisplayHeader = true;
                    }

                    if (!headerIndexes.TryGetValue(columnName, out int columnIndex)) {
                        columnIndex = headers.Count;
                        headers.Add(columnName);
                        headerIndexes.Add(columnName, columnIndex);
                        state.EnsureColumnTypeCapacity(headers.Count);
                    }

                    UpdateObjectExportColumnType(state.InferredColumnTypes, columnIndex, entry.Value);
                }
            }

            state.NormalizeColumnTypeWidth(headers.Count);

            if (headers.Count == 0
                || includeHeaders && state.HasBlankDisplayHeader
                || HasDuplicateObjectExportHeaders(headers)
                || !CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Count)) {
                return false;
            }

            string range = BuildObjectExportRange(startRow, headers.Count, rows.Count, includeHeaders);
            return _excelDocument.RegisterDeferredDirectDictionaryRowsSaveCandidate(
                this,
                Name,
                headers,
                CompleteObjectExportColumnTypes(state.InferredColumnTypes),
                dictionaryRows,
                includeHeaders,
                range);
        }

        private bool TryInsertLegacyDictionaryRowsAsDeferredDirectSave<T>(
            IReadOnlyList<T> rows,
            bool includeHeaders,
            int startRow) {
            var headers = new List<string>();
            var headerIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var dictionaryRows = new System.Collections.IDictionary[rows.Count];
            var state = new FlatDictionaryProjectionState();

            for (int r = 0; r < rows.Count; r++) {
                if (rows[r] is not System.Collections.IDictionary dictionary) {
                    return false;
                }

                dictionaryRows[r] = dictionary;
                foreach (System.Collections.DictionaryEntry entry in dictionary) {
                    object? value = entry.Value;
                    if (!IsFlatDictionaryObjectExportValue(value)) {
                        return false;
                    }

                    string columnName = entry.Key?.ToString() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(columnName)) {
                        state.HasBlankDisplayHeader = true;
                    }

                    if (!headerIndexes.TryGetValue(columnName, out int columnIndex)) {
                        columnIndex = headers.Count;
                        headers.Add(columnName);
                        headerIndexes.Add(columnName, columnIndex);
                        state.EnsureColumnTypeCapacity(headers.Count);
                    }

                    UpdateObjectExportColumnType(state.InferredColumnTypes, columnIndex, value);
                }
            }

            state.NormalizeColumnTypeWidth(headers.Count);

            if (headers.Count == 0
                || includeHeaders && state.HasBlankDisplayHeader
                || HasDuplicateObjectExportHeaders(headers)
                || !CanRegisterDirectTabularSaveCandidate(startRow, 1, headers.Count)) {
                return false;
            }

            string range = BuildObjectExportRange(startRow, headers.Count, rows.Count, includeHeaders);
            return _excelDocument.RegisterDeferredDirectLegacyDictionaryRowsSaveCandidate(
                this,
                Name,
                headers,
                CompleteObjectExportColumnTypes(state.InferredColumnTypes),
                dictionaryRows,
                includeHeaders,
                range);
        }

        private static bool TryProjectFlatDictionaryRow(
            object? item,
            int rowIndex,
            List<string> headers,
            Dictionary<string, int> headerIndexes,
            object?[][] directRows,
            FlatDictionaryProjectionState state) {
            if (item == null) {
                return false;
            }

            object?[] rowValues = new object?[GetFlatDictionaryInitialRowCapacity(item, headers.Count)];
            if (item is Dictionary<string, object?> exactDictionary) {
                foreach (var entry in exactDictionary) {
                    if (!TryAddFlatDictionaryValue(entry.Key, entry.Value)) {
                        return false;
                    }
                }

                directRows[rowIndex] = rowValues;
                return true;
            }

            if (item is IReadOnlyDictionary<string, object?> readOnlyDictionary) {
                foreach (var entry in readOnlyDictionary) {
                    if (!TryAddFlatDictionaryValue(entry.Key, entry.Value)) {
                        return false;
                    }
                }

                directRows[rowIndex] = rowValues;
                return true;
            }

            if (item is IDictionary<string, object?> dictionary) {
                foreach (var entry in dictionary) {
                    if (!TryAddFlatDictionaryValue(entry.Key, entry.Value)) {
                        return false;
                    }
                }

                directRows[rowIndex] = rowValues;
                return true;
            }

            if (item is System.Collections.IDictionary legacyDictionary) {
                foreach (System.Collections.DictionaryEntry entry in legacyDictionary) {
                    string key = entry.Key?.ToString() ?? string.Empty;
                    if (!TryAddFlatDictionaryValue(key, entry.Value)) {
                        return false;
                    }
                }

                directRows[rowIndex] = rowValues;
                return true;
            }

            if (TryProjectPowerShellObjectRow(item, TryAddFlatDictionaryValue)) {
                directRows[rowIndex] = rowValues;
                return true;
            }

            return false;

            bool TryAddFlatDictionaryValue(string? key, object? value) {
                if (!IsFlatDictionaryObjectExportValue(value)) {
                    return false;
                }

                string columnName = key ?? string.Empty;
                if (string.IsNullOrWhiteSpace(columnName)) {
                    state.HasBlankDisplayHeader = true;
                }

                if (!headerIndexes.TryGetValue(columnName, out int columnIndex)) {
                    columnIndex = headers.Count;
                    headers.Add(columnName);
                    headerIndexes.Add(columnName, columnIndex);
                    state.EnsureColumnTypeCapacity(headers.Count);
                    EnsureFlatDictionaryRowCapacity(ref rowValues, headers.Count);
                }

                rowValues[columnIndex] = value;
                UpdateObjectExportColumnType(state.InferredColumnTypes, columnIndex, value);
                return true;
            }
        }

        private static int GetFlatDictionaryInitialRowCapacity(object item, int existingHeaderCount) {
            int entryCount =
                item is System.Collections.ICollection collection
                    ? collection.Count
                    : item is IReadOnlyCollection<KeyValuePair<string, object?>> readOnlyCollection
                        ? readOnlyCollection.Count
                        : 0;

            if (existingHeaderCount == 0) {
                return entryCount;
            }

            return entryCount > existingHeaderCount * 2
                ? existingHeaderCount + entryCount
                : existingHeaderCount;
        }

        private static void EnsureFlatDictionaryRowCapacity(ref object?[] row, int requiredLength) {
            if (row.Length >= requiredLength) {
                return;
            }

            int newLength = row.Length == 0 ? 4 : row.Length * 2;
            if (newLength < requiredLength) {
                newLength = requiredLength;
            }

            Array.Resize(ref row, newLength);
        }

        private static void NormalizeFlatDictionaryRowWidths(object?[][] rows, int columnCount) {
            for (int i = 0; i < rows.Length; i++) {
                if (rows[i].Length == columnCount) {
                    continue;
                }

                object?[] row = rows[i];
                Array.Resize(ref row, columnCount);
                rows[i] = row;
            }
        }

        private sealed class FlatDictionaryProjectionState {
            internal Type?[] InferredColumnTypes = Array.Empty<Type?>();

            internal bool HasBlankDisplayHeader;

            internal void EnsureColumnTypeCapacity(int requiredLength) {
                if (InferredColumnTypes.Length >= requiredLength) {
                    return;
                }

                int newLength = InferredColumnTypes.Length == 0 ? 4 : InferredColumnTypes.Length * 2;
                if (newLength < requiredLength) {
                    newLength = requiredLength;
                }

                Array.Resize(ref InferredColumnTypes, newLength);
            }

            internal void NormalizeColumnTypeWidth(int columnCount) {
                if (InferredColumnTypes.Length == columnCount) {
                    return;
                }

                Array.Resize(ref InferredColumnTypes, columnCount);
            }
        }

        private static bool IsFlatDictionaryObjectExportValue(object? value) {
            return value == null
                || value == DBNull.Value
                || IsObjectExportScalarType(value.GetType());
        }

        private delegate bool TryAddObjectExportValue(string? name, object? value);

        private static bool TryProjectPowerShellObjectRow(object item, TryAddObjectExportValue tryAddValue) {
            Type itemType = item.GetType();
            if (!IsPowerShellObjectExportType(itemType)) {
                return false;
            }

            PowerShellObjectExportPlan plan = PowerShellObjectExportPlans.GetOrAdd(itemType, CreatePowerShellObjectExportPlan);
            if (!plan.CanProject) {
                return false;
            }

            object? propertiesValue;
            try {
                propertiesValue = plan.PropertiesGetter(item);
            } catch {
                return false;
            }

            if (propertiesValue is not IEnumerable properties) {
                return false;
            }

            bool added = false;
            foreach (object? property in properties) {
                if (property == null) {
                    continue;
                }

                Type propertyType = property.GetType();
                PowerShellPropertyExportPlan propertyPlan = plan.GetPropertyPlan(propertyType);
                if (!propertyPlan.CanProject) {
                    continue;
                }

                try {
                    if (propertyPlan.IsGettableGetter != null
                        && propertyPlan.IsGettableGetter(property) is bool isGettable
                        && !isGettable) {
                        continue;
                    }

                    string? name = propertyPlan.NameGetter(property)?.ToString();
                    object? value = propertyPlan.ValueGetter(property);
                    if (!tryAddValue(name, value)) {
                        return false;
                    }

                    added = true;
                } catch {
                    continue;
                }
            }

            return added;
        }

        private static bool IsPowerShellObjectExportType(Type type)
            => string.Equals(type.FullName, "System.Management.Automation.PSObject", StringComparison.Ordinal)
               || string.Equals(type.FullName, "System.Management.Automation.PSCustomObject", StringComparison.Ordinal);

        private static PowerShellObjectExportPlan CreatePowerShellObjectExportPlan(Type type) {
            PropertyInfo? properties = type.GetProperty("Properties", BindingFlags.Public | BindingFlags.Instance);
            if (properties == null || !typeof(IEnumerable).IsAssignableFrom(properties.PropertyType)) {
                return PowerShellObjectExportPlan.NotSupported;
            }

            return new PowerShellObjectExportPlan(CreatePowerShellValueGetter(properties));
        }

        private static PowerShellPropertyExportPlan CreatePowerShellPropertyExportPlan(Type type) {
            PropertyInfo? name = type.GetProperty("Name", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo? value = type.GetProperty("Value", BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo? isGettable = type.GetProperty("IsGettable", BindingFlags.Public | BindingFlags.Instance);
            if (name == null || value == null) {
                return PowerShellPropertyExportPlan.NotSupported;
            }

            return new PowerShellPropertyExportPlan(
                CreatePowerShellValueGetter(name),
                CreatePowerShellValueGetter(value),
                isGettable == null ? null : CreatePowerShellValueGetter(isGettable));
        }

        private static Func<object, object?> CreatePowerShellValueGetter(PropertyInfo property) {
            MethodInfo? getMethod = property.GetMethod;
            if (getMethod == null || property.DeclaringType == null) {
                return row => property.GetValue(row, null);
            }

            try {
                return (Func<object, object?>)CreatePowerShellValueGetterMethod
                    .MakeGenericMethod(property.DeclaringType, property.PropertyType)
                    .Invoke(null, new object[] { getMethod })!;
            } catch {
                return row => property.GetValue(row, null);
            }
        }

        private static readonly MethodInfo CreatePowerShellValueGetterMethod =
            typeof(ExcelSheet).GetMethod(nameof(CreatePowerShellValueGetterCore), BindingFlags.NonPublic | BindingFlags.Static)!;

        private static Func<object, object?> CreatePowerShellValueGetterCore<TTarget, TValue>(MethodInfo getMethod) {
            var getter = (Func<TTarget, TValue>)Delegate.CreateDelegate(typeof(Func<TTarget, TValue>), getMethod);
            return row => getter((TTarget)row!);
        }

        private sealed class PowerShellObjectExportPlan {
            internal static readonly PowerShellObjectExportPlan NotSupported = new();

            private readonly ConcurrentDictionary<Type, PowerShellPropertyExportPlan> _propertyPlans = new();

            private PowerShellObjectExportPlan() {
                PropertiesGetter = _ => null;
                CanProject = false;
            }

            internal PowerShellObjectExportPlan(Func<object, object?> propertiesGetter) {
                PropertiesGetter = propertiesGetter;
                CanProject = true;
            }

            internal bool CanProject { get; }

            internal Func<object, object?> PropertiesGetter { get; }

            internal PowerShellPropertyExportPlan GetPropertyPlan(Type propertyType)
                => _propertyPlans.GetOrAdd(propertyType, CreatePowerShellPropertyExportPlan);
        }

        private sealed class PowerShellPropertyExportPlan {
            internal static readonly PowerShellPropertyExportPlan NotSupported = new();

            private PowerShellPropertyExportPlan() {
                NameGetter = _ => null;
                ValueGetter = _ => null;
                CanProject = false;
            }

            internal PowerShellPropertyExportPlan(
                Func<object, object?> nameGetter,
                Func<object, object?> valueGetter,
                Func<object, object?>? isGettableGetter) {
                NameGetter = nameGetter;
                ValueGetter = valueGetter;
                IsGettableGetter = isGettableGetter;
                CanProject = true;
            }

            internal bool CanProject { get; }

            internal Func<object, object?> NameGetter { get; }

            internal Func<object, object?> ValueGetter { get; }

            internal Func<object, object?>? IsGettableGetter { get; }
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
            if (IsObjectExportScalarType(type)) {
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

