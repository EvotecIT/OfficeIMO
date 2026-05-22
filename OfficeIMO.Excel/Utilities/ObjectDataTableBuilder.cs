using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using OfficeIMO.Shared;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Builds DataTable instances from object sequences by projecting properties or dictionary keys into columns.
    /// </summary>
    public static class ObjectDataTableBuilder {
        /// <summary>
        /// Creates a DataTable from objects by using dictionary keys or public readable properties.
        /// </summary>
        /// <param name="items">Sequence of objects to convert into table rows.</param>
        /// <param name="tableName">Optional DataTable name.</param>
        /// <returns>Populated DataTable.</returns>
        [RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly or pre-flatten items.")]
        public static DataTable FromObjects(IEnumerable<object?> items, string tableName = "Data") {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            IReadOnlyList<object?> rows = items as IReadOnlyList<object?> ?? items.ToList();
            if (rows.Count == 0) {
                throw new ArgumentException("Provide at least one data row.", nameof(items));
            }

            var first = rows[0];
            if (first == null) {
                throw new ArgumentException("Data rows cannot be null.", nameof(items));
            }

            var columnNames = ObjectDataHelpers.GetColumnNames(first);
            if (columnNames.Count == 0) {
                throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries.");
            }

            var table = new DataTable(tableName);
            foreach (var name in columnNames) {
                table.Columns.Add(name, typeof(object));
            }

            var projector = ObjectRowProjector.Create(first, columnNames);
            table.MinimumCapacity = Math.Max(table.MinimumCapacity, rows.Count);
            var values = projector.CreateValuesBuffer();
            table.BeginLoadData();
            try {
                for (int i = 0; i < rows.Count; i++) {
                    var item = rows[i];
                    if (item == null) {
                        throw new InvalidOperationException("Data rows cannot contain null entries.");
                    }

                    projector.FillValues(item, values);
                    table.Rows.Add(values);
                }
            } finally {
                table.EndLoadData();
            }

            return table;
        }

        private sealed class ObjectRowProjector {
            private readonly string[] _columns;
            private readonly PropertyInfo[]? _properties;
            private readonly Type? _sourceType;

            private ObjectRowProjector(IReadOnlyList<string> columns, PropertyInfo[]? properties, Type? sourceType) {
                _columns = columns.ToArray();
                _properties = properties;
                _sourceType = sourceType;
            }

            internal static ObjectRowProjector Create(object first, IReadOnlyList<string> columns) {
                var firstType = first.GetType();
                if (IsDictionaryLike(first)) {
                    return new ObjectRowProjector(columns, properties: null, sourceType: null);
                }

                var properties = new PropertyInfo[columns.Count];
                for (int i = 0; i < columns.Count; i++) {
                    var property = firstType.GetProperty(columns[i], BindingFlags.Public | BindingFlags.Instance);
                    if (property == null || !property.CanRead || property.GetIndexParameters().Length != 0) {
                        return new ObjectRowProjector(columns, properties: null, sourceType: null);
                    }

                    properties[i] = property;
                }

                return new ObjectRowProjector(columns, properties, firstType);
            }

            internal object?[] CreateValuesBuffer() => new object?[_columns.Length];

            internal void FillValues(object item, object?[] values) {
                if (TryFillDictionaryValues(item, values)) {
                    return;
                }

                if (_properties != null && item.GetType() == _sourceType) {
                    for (int i = 0; i < _properties.Length; i++) {
                        values[i] = _properties[i].GetValue(item) ?? DBNull.Value;
                    }

                    return;
                }

                for (int i = 0; i < _columns.Length; i++) {
                    values[i] = ObjectDataHelpers.GetValue(item, _columns[i]) ?? DBNull.Value;
                }
            }

            private bool TryFillDictionaryValues(object item, object?[] values) {
                if (item is Dictionary<string, object?> exactDictionary) {
                    for (int i = 0; i < _columns.Length; i++) {
                        values[i] = exactDictionary.TryGetValue(_columns[i], out var value) ? value ?? DBNull.Value : DBNull.Value;
                    }

                    return true;
                }

                if (item is IReadOnlyDictionary<string, object?> readOnlyDictionary) {
                    for (int i = 0; i < _columns.Length; i++) {
                        values[i] = readOnlyDictionary.TryGetValue(_columns[i], out var value) ? value ?? DBNull.Value : DBNull.Value;
                    }

                    return true;
                }

                if (item is IDictionary<string, object?> dictionary) {
                    for (int i = 0; i < _columns.Length; i++) {
                        values[i] = dictionary.TryGetValue(_columns[i], out var value) ? value ?? DBNull.Value : DBNull.Value;
                    }

                    return true;
                }

                if (item is System.Collections.IDictionary legacyDictionary) {
                    for (int i = 0; i < _columns.Length; i++) {
                        values[i] = GetLegacyDictionaryValue(legacyDictionary, _columns[i]) ?? DBNull.Value;
                    }

                    return true;
                }

                return false;
            }

            private static object? GetLegacyDictionaryValue(System.Collections.IDictionary dictionary, string column) {
                if (dictionary.Contains(column)) {
                    return dictionary[column];
                }

                foreach (System.Collections.DictionaryEntry entry in dictionary) {
                    var key = entry.Key?.ToString();
                    if (string.Equals(key, column, StringComparison.OrdinalIgnoreCase)) {
                        return entry.Value;
                    }
                }

                return null;
            }

            private static bool IsDictionaryLike(object item) {
                return item is Dictionary<string, object?>
                    || item is IReadOnlyDictionary<string, object?>
                    || item is IDictionary<string, object?>
                    || item is System.Collections.IDictionary;
            }
        }

    }
}
