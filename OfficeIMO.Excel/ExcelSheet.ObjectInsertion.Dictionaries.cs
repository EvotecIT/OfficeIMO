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

            if (TryInsertPowerShellObjectRowsAsCellValues(rows, includeHeaders, startRow)) {
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
            var headers = new List<string>();
            var headerIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
            var directRows = new object?[rows.Count][];
            var state = new FlatDictionaryProjectionState();

            for (int r = 0; r < rows.Count; r++) {
                if (rows[r] is not Dictionary<string, object?> dictionary
                    || !TryProjectExactDictionaryRow(dictionary, r, headers, headerIndexes, directRows, state)) {
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

            string range = BuildObjectExportRange(startRow, headers.Count, rows.Count, includeHeaders);
            return TryInsertRowsAsDeferredDirectSave(
                Name,
                headers,
                CompleteObjectExportColumnTypes(state.InferredColumnTypes),
                directRows,
                startRow,
                includeHeaders,
                range);
        }

        private static bool TryProjectExactDictionaryRow(
            Dictionary<string, object?> dictionary,
            int rowIndex,
            List<string> headers,
            Dictionary<string, int> headerIndexes,
            object?[][] directRows,
            FlatDictionaryProjectionState state) {
            var rowValues = new object?[Math.Max(dictionary.Count, headers.Count)];
            foreach (var entry in dictionary) {
                object? value = entry.Value;
                if (!IsFlatDictionaryObjectExportValue(value)) {
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
                    EnsureFlatDictionaryRowCapacity(ref rowValues, headers.Count);
                }

                rowValues[columnIndex] = value;
                UpdateObjectExportColumnType(state.InferredColumnTypes, columnIndex, value);
            }

            directRows[rowIndex] = rowValues;
            return true;
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
            var headerIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
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
                || HasDuplicateObjectExportHeaders(headers, StringComparer.Ordinal)
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

        private static bool TryGetFlatObjectExportValueType(object? value, out Type? valueType) {
            switch (value) {
                case null:
                case DBNull _:
                    valueType = null;
                    return true;
                case string _:
                    valueType = typeof(string);
                    return true;
                case int _:
                    valueType = typeof(int);
                    return true;
                case bool _:
                    valueType = typeof(bool);
                    return true;
                case DateTime _:
                    valueType = typeof(DateTime);
                    return true;
                case double _:
                    valueType = typeof(double);
                    return true;
                case long _:
                    valueType = typeof(long);
                    return true;
                case decimal _:
                    valueType = typeof(decimal);
                    return true;
                case float _:
                    valueType = typeof(float);
                    return true;
                case DateTimeOffset _:
                    valueType = typeof(DateTimeOffset);
                    return true;
                case TimeSpan _:
                    valueType = typeof(TimeSpan);
                    return true;
                case Guid _:
                    valueType = typeof(Guid);
                    return true;
#if NET6_0_OR_GREATER
                case DateOnly _:
                    valueType = typeof(DateOnly);
                    return true;
                case TimeOnly _:
                    valueType = typeof(TimeOnly);
                    return true;
#endif
            }

            Type type = value.GetType();
            if (IsObjectExportScalarType(type)) {
                valueType = type;
                return true;
            }

            valueType = null;
            return false;
        }

        private delegate bool TryAddObjectExportValue(string? name, object? value);
        private delegate bool TryAddIndexedObjectExportValue(int propertyIndex, string? name, object? value);


        [RequiresUnreferencedCode("Runtime-object flattening is a compatibility path. Use InsertObjects with explicit column selectors in NativeAOT applications.")]
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
    }
}
