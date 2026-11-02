using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Data;
using System.Reflection;
using System.Text;
using OfficeIMO.Data;
using OfficeIMO.Spreadsheet;
using OfficeColor = OfficeIMO.OfficeColor;

namespace OfficeIMO.Excel;

internal sealed partial class ExcelObjectTableProjection {
        private static string TransformHeader(string path, ObjectFlattenerOptions opts) {
            if (opts.TableOptions is { } tableOptions && (tableOptions.HeaderRenames.ContainsKey(path) || tableOptions.HeaderTransform != null))
                return OfficeIMO.Core.Internal.ObjectDataHelpers.GetTableHeader(path, tableOptions);
            path = opts.GetHeaderPath(path);
            return opts.HeaderCase switch {
                HeaderCase.Pascal => TransformHeaderPascal(path),
                HeaderCase.Title => string.Join(" ", path.Split('.').Select(s => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLowerInvariant()))),
                _ => path
            };
        }

        private static List<string> BuildTransformedHeaders(IReadOnlyList<string> paths, ObjectFlattenerOptions options) {
            var headers = new List<string>(paths.Count);
            for (int i = 0; i < paths.Count; i++) {
                headers.Add(TransformHeader(paths[i], options));
            }

            return headers;
        }

        private static string TransformHeaderPascal(string path) {
            if (path.Length == 0) {
                ThrowEmptyHeaderSegment();
            }

            var builder = new StringBuilder(path.Length);
            int segmentStart = 0;
            for (int i = 0; i <= path.Length; i++) {
                if (i < path.Length && path[i] != '.') {
                    continue;
                }

                AppendPascalSegment(builder, path, segmentStart, i - segmentStart);
                segmentStart = i + 1;
            }

            return builder.ToString();
        }

        private static void AppendPascalSegment(StringBuilder builder, string path, int start, int length) {
            if (length == 0) {
                ThrowEmptyHeaderSegment();
            }

            builder.Append(char.ToUpperInvariant(path[start]));
            if (length > 1) {
                builder.Append(path, start + 1, length - 1);
            }
        }

        private static void ThrowEmptyHeaderSegment() => throw new IndexOutOfRangeException();

        private static bool CanUseRowsFromDataTable(IReadOnlyList<string> headers) {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string header in headers) {
                if (string.IsNullOrWhiteSpace(header)) {
                    return false;
                }

                if (!seen.Add(header)) {
                    return false;
                }
            }

            return headers.Count > 0;
        }

        [UnconditionalSuppressMessage("Trimming", "IL2062", Justification = "AddTable limits inferred DataColumn types to supported scalar framework types; custom runtime types fall back to object/string cell storage.")]
        private static DataTable CreateRowsFromDataTable(IReadOnlyList<string> headers, IReadOnlyList<object?[]> rowValues) {
            var table = new DataTable("AddTable") {
                Locale = CultureInfo.InvariantCulture
            };

            var columnTypes = InferRowsFromColumnTypes(rowValues, headers.Count);
            for (int i = 0; i < headers.Count; i++) {
                table.Columns.Add(headers[i], columnTypes[i]);
            }

            table.BeginLoadData();
            try {
                foreach (object?[] values in rowValues) {
                    var row = new object[headers.Count];
                    for (int i = 0; i < headers.Count; i++) {
                        row[i] = CoerceRowsFromDataTableValue(i < values.Length ? values[i] : null, columnTypes[i]);
                    }

                    table.Rows.Add(row);
                }
            } finally {
                table.EndLoadData();
            }

            return table;
        }

        private static RowsFromSimpleTypePlan GetRowsFromSimpleTypePlan(Type type)
            => RowsFromSimpleTypePlans.GetOrAdd(type, CreateRowsFromSimpleTypePlan);

        [UnconditionalSuppressMessage("Trimming", "IL2070", Justification = "RowsFrom<T> preserves public properties on T; this cache only discovers those public row properties.")]
        private static RowsFromSimpleTypePlan CreateRowsFromSimpleTypePlan(Type type) {
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(property => property.GetIndexParameters().Length == 0 && property.GetMethod != null)
                .OrderBy(ObjectFlattener.GetMetadataOrder)
                .ThenBy(property => property.Name, StringComparer.Ordinal)
                .ToArray();
            if (properties.Length == 0 || properties.Any(property => !IsRowsFromDirectSaveScalarType(property.PropertyType))) {
                return RowsFromSimpleTypePlan.NotSupported;
            }

            var headers = new string[properties.Length];
            for (int i = 0; i < properties.Length; i++) {
                headers[i] = properties[i].Name;
            }

            var staticColumnTypes = new Type[properties.Length];
            var staticBlankAsEmptyString = new bool[properties.Length];
            var staticRequiresBlankCoercion = new bool[properties.Length];
            var inferenceFallbackTypes = new Type[properties.Length];
            var inferColumns = new bool[properties.Length];
            var getters = new RowsFromSimpleValueGetter[properties.Length];
            var inferenceColumnIndexes = new List<int>();
            bool hasInferenceColumns = false;
            for (int i = 0; i < properties.Length; i++) {
                getters[i] = CreateRowsFromSimpleValueGetter(properties[i]);
                Type propertyType = properties[i].PropertyType;
                Type? nullableUnderlyingType = Nullable.GetUnderlyingType(propertyType);
                Type declaredType = nullableUnderlyingType ?? propertyType;
                if (declaredType == typeof(string)) {
                    staticColumnTypes[i] = typeof(string);
                    staticBlankAsEmptyString[i] = true;
                    staticRequiresBlankCoercion[i] = true;
                    inferenceFallbackTypes[i] = typeof(string);
                } else if (nullableUnderlyingType != null && CanUseStaticNullableRowsFromColumnType(nullableUnderlyingType)) {
                    staticColumnTypes[i] = nullableUnderlyingType;
                    staticBlankAsEmptyString[i] = false;
                    staticRequiresBlankCoercion[i] = true;
                    inferenceFallbackTypes[i] = nullableUnderlyingType;
                } else if (!declaredType.IsValueType || nullableUnderlyingType != null) {
                    inferColumns[i] = true;
                    inferenceColumnIndexes.Add(i);
                    hasInferenceColumns = true;
                    inferenceFallbackTypes[i] = typeof(object);
                } else {
                    staticColumnTypes[i] = declaredType;
                    staticBlankAsEmptyString[i] = false;
                    staticRequiresBlankCoercion[i] = false;
                    inferenceFallbackTypes[i] = declaredType;
                }
            }

            return CanUseRowsFromDataTable(headers)
                ? new RowsFromSimpleTypePlan(getters, headers, staticColumnTypes, staticBlankAsEmptyString, staticRequiresBlankCoercion, inferenceFallbackTypes, inferColumns, inferenceColumnIndexes.ToArray(), hasInferenceColumns, canUseDirectSave: true)
                : RowsFromSimpleTypePlan.NotSupported;
        }

        private static RowsFromSimpleValueGetter CreateRowsFromSimpleValueGetter(PropertyInfo property) {
            return property.GetValue;
        }

        private static bool IsRowsFromDirectSaveScalarType(Type type) {
            type = Nullable.GetUnderlyingType(type) ?? type;
            if (type == typeof(string)) {
                return true;
            }

            if (typeof(System.Collections.IEnumerable).IsAssignableFrom(type)) {
                return false;
            }

            return type.IsPrimitive
                || type.IsEnum
                || type == typeof(decimal)
                || type == typeof(DateTime)
                || type == typeof(DateTimeOffset)
                || type == typeof(TimeSpan)
                || type == typeof(Guid);
        }

        private static bool CanUseStaticNullableRowsFromColumnType(Type type) {
            if (type == typeof(DateTime)
                || type == typeof(DateTimeOffset)
                || type == typeof(TimeSpan)) {
                return false;
            }

#if NET6_0_OR_GREATER
            if (type == typeof(DateOnly) || type == typeof(TimeOnly)) {
                return false;
            }
#endif

            return type.IsPrimitive
                || type.IsEnum
                || type == typeof(decimal)
                || type == typeof(Guid);
        }

        private static object?[][] MaterializeSimpleRowsFromProperties<T>(RowsFromSimpleTypePlan typePlan, IReadOnlyList<T> rows, out Type[] columnTypes) {
            var directRows = new object?[rows.Count][];
            if (!typePlan.HasInferenceColumns) {
                MaterializeStaticSimpleRowsFromProperties(typePlan, rows, directRows);
                columnTypes = typePlan.StaticColumnTypes;
                return directRows;
            }

            columnTypes = InferRowsFromPropertyColumnTypes(typePlan, rows, directRows);
            int[] inferenceColumnIndexes = typePlan.InferenceColumnIndexes;
            for (int row = 0; row < directRows.Length; row++) {
                object?[] values = directRows[row];
                for (int i = 0; i < inferenceColumnIndexes.Length; i++) {
                    int column = inferenceColumnIndexes[i];
                    values[column] = CoerceRowsFromDataTableValue(values[column], columnTypes[column]);
                }
            }

            return directRows;
        }

        private static void MaterializeStaticSimpleRowsFromProperties<T>(RowsFromSimpleTypePlan typePlan, IReadOnlyList<T> rows, object?[][] directRows) {
            RowsFromSimpleValueGetter[] getters = typePlan.Getters;
            bool[] blankAsEmptyString = typePlan.StaticBlankAsEmptyString;
            bool[] requiresBlankCoercion = typePlan.StaticRequiresBlankCoercion;
            int columnCount = getters.Length;
            for (int row = 0; row < rows.Count; row++) {
                T sourceRow = rows[row];
                var values = new object?[columnCount];
                for (int column = 0; column < columnCount; column++) {
                    object? value = getters[column](sourceRow);
                    values[column] = requiresBlankCoercion[column]
                        ? CoerceRowsFromDirectSaveValue(value, blankAsEmptyString[column])
                        : value;
                }

                directRows[row] = values;
            }
        }

        private static Type[] InferRowsFromPropertyColumnTypes<T>(RowsFromSimpleTypePlan typePlan, IReadOnlyList<T> rows, object?[][] directRows) {
            RowsFromSimpleValueGetter[] getters = typePlan.Getters;
            bool[] inferColumns = typePlan.InferColumns;
            bool[] staticBlankAsEmptyString = typePlan.StaticBlankAsEmptyString;
            bool[] staticRequiresBlankCoercion = typePlan.StaticRequiresBlankCoercion;
            int columnCount = getters.Length;
            var columnTypes = new Type[columnCount];
            Array.Copy(typePlan.StaticColumnTypes, columnTypes, columnTypes.Length);
            Type?[]? inferredTypes = typePlan.HasInferenceColumns ? new Type?[columnCount] : null;

            for (int row = 0; row < rows.Count; row++) {
                T sourceRow = rows[row];
                var values = new object?[columnCount];
                for (int column = 0; column < columnCount; column++) {
                    object? value = getters[column](sourceRow);
                    if (!inferColumns[column]) {
                        values[column] = staticRequiresBlankCoercion[column]
                            ? CoerceRowsFromDirectSaveValue(value, staticBlankAsEmptyString[column])
                            : value;
                        continue;
                    }

                    values[column] = value;
                    if (IsRowsFromBlankValue(value)) {
                        continue;
                    }

                    Type valueType = value!.GetType();
                    Type? inferred = inferredTypes![column];
                    if (inferred == null) {
                        inferredTypes[column] = valueType;
                    } else if (inferred != valueType) {
                        inferredTypes[column] = typeof(object);
                    }
                }

                directRows[row] = values;
            }

            int[] inferenceColumnIndexes = typePlan.InferenceColumnIndexes;
            for (int i = 0; i < inferenceColumnIndexes.Length; i++) {
                int column = inferenceColumnIndexes[i];
                columnTypes[column] = inferredTypes![column] ?? typePlan.InferenceFallbackTypes[column];
            }

            return columnTypes;
        }

        private static Type[] InferRowsFromColumnTypes(IReadOnlyList<object?[]> rowValues, int columnCount) {
            var columnTypes = new Type[columnCount];
            for (int column = 0; column < columnCount; column++) {
                Type? inferred = null;
                for (int row = 0; row < rowValues.Count; row++) {
                    object? value = column < rowValues[row].Length ? rowValues[row][column] : null;
                    if (IsRowsFromBlankValue(value)) {
                        continue;
                    }

                    Type valueType = value!.GetType();
                    if (inferred == null) {
                        inferred = valueType;
                        continue;
                    }

                    if (inferred != valueType) {
                        inferred = typeof(object);
                        break;
                    }
                }

                columnTypes[column] = inferred ?? typeof(string);
            }

            return columnTypes;
        }

        private static bool IsRowsFromBlankValue(object? value) {
            return value == null || value == DBNull.Value || value is string text && text.Length == 0;
        }

        private static object CoerceRowsFromDataTableValue(object? value, Type columnType) {
            if (IsRowsFromBlankValue(value)) {
                return columnType == typeof(string) || columnType == typeof(object)
                    ? string.Empty
                    : DBNull.Value;
            }

            return value!;
        }

        private static object CoerceRowsFromDirectSaveValue(object? value, bool blankAsEmptyString) {
            if (IsRowsFromBlankValue(value)) {
                return blankAsEmptyString ? string.Empty : DBNull.Value;
            }

            return value!;
        }

        private void AddRowsFromCellValues(List<(int Row, int Column, object Value)> cells, int startRow, IReadOnlyList<string> headers, IReadOnlyList<object?[]> rowValues) {
            for (int i = 0; _includeHeaders && i < headers.Count; i++) {
                cells.Add((startRow, i + _startColumn, headers[i]));
            }

            for (int r = 0; r < rowValues.Count; r++) {
                object?[] values = rowValues[r];
                for (int c = 0; c < headers.Count; c++) {
                    object? value = c < values.Length ? values[c] : null;
                    cells.Add((startRow + r + HeaderRows, c + _startColumn, value ?? string.Empty));
                }
            }
        }

        private void AddSimpleRowsFromCellValues(int startRow, IReadOnlyList<string> headers, IReadOnlyList<object?[]> rowValues) {
            int totalCellCount = checked((rowValues.Count + HeaderRows) * headers.Count);
            var cells = new (int Row, int Column, object Value)[totalCellCount];
            int cellIndex = 0;
            for (int i = 0; _includeHeaders && i < headers.Count; i++) {
                cells[cellIndex++] = (startRow, i + _startColumn, headers[i]);
            }

            for (int r = 0; r < rowValues.Count; r++) {
                object?[] values = rowValues[r];
                for (int c = 0; c < headers.Count; c++) {
                    cells[cellIndex++] = (startRow + r + HeaderRows, c + _startColumn, values[c] ?? string.Empty);
                }
            }

            Sheet!.CellValues(cells);
        }

        private static string ColumnLetter(int column) {
            var dividend = column;
            var columnName = string.Empty;
            while (dividend > 0) {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }


}
