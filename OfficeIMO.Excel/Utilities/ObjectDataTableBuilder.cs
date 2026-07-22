using System.Collections.Concurrent;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using OfficeIMO.Drawing.Internal;

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
        /// <param name="options">Optional value-normalization settings for projected cell values.</param>
        /// <returns>Populated DataTable.</returns>
        [RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly or pre-flatten items.")]
        public static DataTable FromObjects(IEnumerable<object?> items, string tableName = "Data", ObjectDataTableBuilderOptions? options = null) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            options ??= new ObjectDataTableBuilderOptions();
            IReadOnlyList<object?> rows = items as IReadOnlyList<object?> ?? items.ToList();
            if (rows.Count == 0) {
                throw new ArgumentException("Provide at least one data row.", nameof(items));
            }

            var first = rows[0];
            if (first == null) {
                throw new ArgumentException("Data rows cannot be null.", nameof(items));
            }

            if (IsScalarValue(first)) {
                throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries.");
            }

            var columnNames = ObjectDataHelpers.GetColumnNames(first);
            if (columnNames.Count == 0 || !columnNames.Any(name => !string.IsNullOrWhiteSpace(name))) {
                throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries.");
            }

            var table = new DataTable(tableName);
            foreach (var name in columnNames) {
                table.Columns.Add(name, typeof(object));
            }

            var projector = ObjectRowProjector.Create(first, columnNames, options);
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
                    projector.ResetValues(values);
                }
            } finally {
                table.EndLoadData();
            }

            return table;
        }

        private static bool IsScalarValue(object item) {
            Type type = item.GetType();
            return type.IsPrimitive ||
                   type.IsEnum ||
                   item is string or decimal or DateTime or DateTimeOffset or TimeSpan or Guid;
        }

        [RequiresUnreferencedCode("This projector belongs to the runtime-object compatibility API. Use typed selectors or pre-flattened dictionaries in NativeAOT applications.")]
        private sealed class ObjectRowProjector {
            private static readonly ConcurrentDictionary<Type, CachedObjectRowProjection> TypedProjectionCache = new();

            private readonly string[] _columns;
            private readonly Dictionary<string, int>? _ordinalColumnIndexes;
            private readonly Dictionary<string, int>? _ignoreCaseColumnIndexes;
            private readonly ObjectValueGetter[]? _getters;
            private readonly Type? _sourceType;
            private readonly ObjectDataTableBuilderOptions _options;
            private int[]? _sparseTouchedColumns;
            private byte[]? _sparseColumnStates;
            private int _sparseTouchedCount;
            private bool _valuesReadyForSparseProjection;

            private ObjectRowProjector(string[] columns, ObjectValueGetter[]? getters, Type? sourceType, bool createOrdinalColumnIndexes, ObjectDataTableBuilderOptions options) {
                _columns = columns;
                _ordinalColumnIndexes = createOrdinalColumnIndexes ? CreateOrdinalColumnIndexes(columns) : null;
                _ignoreCaseColumnIndexes = createOrdinalColumnIndexes ? CreateIgnoreCaseColumnIndexes(columns) : null;
                _getters = getters;
                _sourceType = sourceType;
                _options = options;
            }

            internal static ObjectRowProjector Create(object first, IReadOnlyList<string> columns, ObjectDataTableBuilderOptions options) {
                var firstType = first.GetType();
                if (IsDictionaryLike(first)) {
                    return new ObjectRowProjector(columns.ToArray(), getters: null, sourceType: null, createOrdinalColumnIndexes: true, options);
                }

                var cachedProjection = TypedProjectionCache.GetOrAdd(firstType, type => CreateCachedProjection(type, columns));
                if (ColumnsMatch(cachedProjection.Columns, columns)) {
                    return new ObjectRowProjector(cachedProjection.Columns, cachedProjection.Getters, firstType, createOrdinalColumnIndexes: false, options);
                }

                var getters = CreateGetters(firstType, columns);
                if (getters == null) {
                    return new ObjectRowProjector(columns.ToArray(), getters: null, sourceType: null, createOrdinalColumnIndexes: false, options);
                }

                return new ObjectRowProjector(columns.ToArray(), getters, firstType, createOrdinalColumnIndexes: false, options);
            }

            private static CachedObjectRowProjection CreateCachedProjection(Type firstType, IReadOnlyList<string> columns) {
                var getters = CreateGetters(firstType, columns);
                return new CachedObjectRowProjection(columns.ToArray(), getters);
            }

            private static ObjectValueGetter[]? CreateGetters(Type firstType, IReadOnlyList<string> columns) {
                var getters = new ObjectValueGetter[columns.Count];
                for (int i = 0; i < columns.Count; i++) {
                    var property = firstType.GetProperty(columns[i], BindingFlags.Public | BindingFlags.Instance);
                    if (property == null || !property.CanRead || property.GetIndexParameters().Length != 0) {
                        return null;
                    }

                    getters[i] = CreateObjectValueGetter(property);
                }

                return getters;
            }

            private static bool ColumnsMatch(string[] cachedColumns, IReadOnlyList<string> columns) {
                if (cachedColumns.Length != columns.Count) {
                    return false;
                }

                for (int i = 0; i < cachedColumns.Length; i++) {
                    if (!string.Equals(cachedColumns[i], columns[i], StringComparison.Ordinal)) {
                        return false;
                    }
                }

                return true;
            }

            internal object?[] CreateValuesBuffer() => new object?[_columns.Length];

            internal void FillValues(object item, object?[] values) {
                _sparseTouchedCount = 0;
                if (TryFillDictionaryValues(item, values)) {
                    return;
                }

                if (_getters != null && item.GetType() == _sourceType) {
                    for (int i = 0; i < _getters.Length; i++) {
                        values[i] = NormalizeCellValue(_getters[i](item), _options) ?? DBNull.Value;
                    }

                    _valuesReadyForSparseProjection = false;
                    return;
                }

                for (int i = 0; i < _columns.Length; i++) {
                    values[i] = NormalizeCellValue(ObjectDataHelpers.GetValue(item, _columns[i]), _options) ?? DBNull.Value;
                }

                _valuesReadyForSparseProjection = false;
            }

            internal void ResetValues(object?[] values) {
                if (_sparseTouchedCount == 0) {
                    return;
                }

                for (int i = 0; i < _sparseTouchedCount; i++) {
                    int columnIndex = _sparseTouchedColumns![i];
                    values[columnIndex] = DBNull.Value;
                    if (_sparseColumnStates != null) {
                        _sparseColumnStates[columnIndex] = 0;
                    }
                }

                _sparseTouchedCount = 0;
            }

            private bool TryFillDictionaryValues(object item, object?[] values) {
                if (item is Dictionary<string, object?> exactDictionary) {
                    if (CanUseOrdinalEntryProjection(exactDictionary.Comparer)
                        && TryFillSparseDictionaryValues(exactDictionary, values)) {
                        return true;
                    }

                    for (int i = 0; i < _columns.Length; i++) {
                        values[i] = exactDictionary.TryGetValue(_columns[i], out var value) ? NormalizeCellValue(value, _options) ?? DBNull.Value : DBNull.Value;
                    }

                    _valuesReadyForSparseProjection = false;
                    return true;
                }

                if (item is IReadOnlyDictionary<string, object?> readOnlyDictionary) {
                    for (int i = 0; i < _columns.Length; i++) {
                        values[i] = readOnlyDictionary.TryGetValue(_columns[i], out var value) ? NormalizeCellValue(value, _options) ?? DBNull.Value : DBNull.Value;
                    }

                    _valuesReadyForSparseProjection = false;
                    return true;
                }

                if (item is IDictionary<string, object?> dictionary) {
                    for (int i = 0; i < _columns.Length; i++) {
                        values[i] = dictionary.TryGetValue(_columns[i], out var value) ? NormalizeCellValue(value, _options) ?? DBNull.Value : DBNull.Value;
                    }

                    _valuesReadyForSparseProjection = false;
                    return true;
                }

                if (item is System.Collections.IDictionary legacyDictionary) {
                    if (TryFillSparseLegacyDictionaryValues(legacyDictionary, values)) {
                        return true;
                    }

                    for (int i = 0; i < _columns.Length; i++) {
                        values[i] = NormalizeCellValue(GetLegacyDictionaryValue(legacyDictionary, _columns[i]), _options) ?? DBNull.Value;
                    }

                    _valuesReadyForSparseProjection = false;
                    return true;
                }

                return false;
            }

            private bool TryFillSparseDictionaryValues(IEnumerable<KeyValuePair<string, object?>> dictionary, object?[] values) {
                if (_ordinalColumnIndexes == null || dictionary is ICollection<KeyValuePair<string, object?>> collection && collection.Count >= _columns.Length) {
                    return false;
                }

                if (!_valuesReadyForSparseProjection) {
                    FillDbNull(values);
                    _valuesReadyForSparseProjection = true;
                }

                foreach (var entry in dictionary) {
                    if (_ordinalColumnIndexes.TryGetValue(entry.Key, out int columnIndex)) {
                        values[columnIndex] = NormalizeCellValue(entry.Value, _options) ?? DBNull.Value;
                        TrackSparseTouchedColumn(columnIndex);
                    }
                }

                return true;
            }

            private bool TryFillSparseLegacyDictionaryValues(System.Collections.IDictionary dictionary, object?[] values) {
                if (_ordinalColumnIndexes == null
                    || _ignoreCaseColumnIndexes == null
                    || dictionary.Count >= _columns.Length) {
                    return false;
                }

                if (!_valuesReadyForSparseProjection) {
                    FillDbNull(values);
                    _valuesReadyForSparseProjection = true;
                }

                _sparseColumnStates ??= new byte[_columns.Length];
                foreach (System.Collections.DictionaryEntry entry in dictionary) {
                    string key = entry.Key?.ToString() ?? string.Empty;
                    if (_ordinalColumnIndexes.TryGetValue(key, out int exactColumnIndex)) {
                        TrackSparseLegacyColumn(exactColumnIndex);
                        values[exactColumnIndex] = NormalizeCellValue(entry.Value, _options) ?? DBNull.Value;
                        _sparseColumnStates[exactColumnIndex] = 2;
                        continue;
                    }

                    if (_ignoreCaseColumnIndexes.TryGetValue(key, out int ignoreCaseColumnIndex)
                        && _sparseColumnStates[ignoreCaseColumnIndex] == 0) {
                        TrackSparseLegacyColumn(ignoreCaseColumnIndex);
                        values[ignoreCaseColumnIndex] = NormalizeCellValue(entry.Value, _options) ?? DBNull.Value;
                        _sparseColumnStates[ignoreCaseColumnIndex] = 1;
                    }
                }

                return true;
            }

            private void TrackSparseTouchedColumn(int columnIndex) {
                _sparseTouchedColumns ??= new int[_columns.Length];
                _sparseTouchedColumns[_sparseTouchedCount++] = columnIndex;
            }

            private void TrackSparseLegacyColumn(int columnIndex) {
                if (_sparseColumnStates![columnIndex] != 0) {
                    return;
                }

                TrackSparseTouchedColumn(columnIndex);
            }

            private static Dictionary<string, int> CreateOrdinalColumnIndexes(string[] columns) {
                var indexes = new Dictionary<string, int>(columns.Length, StringComparer.Ordinal);
                for (int i = 0; i < columns.Length; i++) {
                    indexes[columns[i]] = i;
                }

                return indexes;
            }

            private static Dictionary<string, int> CreateIgnoreCaseColumnIndexes(string[] columns) {
                var indexes = new Dictionary<string, int>(columns.Length, StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < columns.Length; i++) {
                    if (!indexes.ContainsKey(columns[i])) {
                        indexes.Add(columns[i], i);
                    }
                }

                return indexes;
            }

            private static void FillDbNull(object?[] values) {
                for (int i = 0; i < values.Length; i++) {
                    values[i] = DBNull.Value;
                }
            }

            private static bool CanUseOrdinalEntryProjection(IEqualityComparer<string> comparer) {
                return ReferenceEquals(comparer, EqualityComparer<string>.Default)
                    || ReferenceEquals(comparer, StringComparer.Ordinal);
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

            private static object? NormalizeCellValue(object? value, ObjectDataTableBuilderOptions options) {
                if (value == null || value == DBNull.Value || !options.NormalizeCollectionValues) {
                    return value;
                }

                if (value is string || value is System.Collections.IDictionary) {
                    return value;
                }

                if (value is IEnumerable enumerable) {
                    return JoinCollection(enumerable, options.CollectionSeparator);
                }

                return value;
            }

            private static string JoinCollection(IEnumerable values, string separator) {
                var parts = new List<string>();
                foreach (var value in values) {
                    parts.Add(value?.ToString() ?? string.Empty);
                }

                return string.Join(separator ?? string.Empty, parts);
            }

            private static ObjectValueGetter CreateObjectValueGetter(PropertyInfo property)
                => row => property.GetValue(row, null);

            private sealed class CachedObjectRowProjection {
                internal CachedObjectRowProjection(string[] columns, ObjectValueGetter[]? getters) {
                    Columns = columns;
                    Getters = getters;
                }

                internal string[] Columns { get; }

                internal ObjectValueGetter[]? Getters { get; }
            }
        }

        private delegate object? ObjectValueGetter(object row);

    }
}
