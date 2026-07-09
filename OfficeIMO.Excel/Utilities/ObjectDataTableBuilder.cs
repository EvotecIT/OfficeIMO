using System.Collections;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using OfficeIMO.Data;

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
            var rows = AsReadOnlyList(items);
            ValidateRows(rows);
            ValidateFirstRowShape(rows[0]!);

            return TabularDataTableBuilder.FromItems(rows, new TabularDataOptions {
                TableName = tableName,
                ColumnDiscoveryMode = TabularColumnDiscoveryMode.FirstRow,
                NormalizeValue = value => NormalizeCellValue(value, options)
            });
        }

        private static IReadOnlyList<object?> AsReadOnlyList(IEnumerable<object?> items) {
            if (items is IReadOnlyList<object?> list) {
                return list;
            }

            return items.ToList();
        }

        private static void ValidateRows(IReadOnlyList<object?> items) {
            if (items.Count == 0) {
                throw new ArgumentException("Provide at least one data row.", nameof(items));
            }

            if (items[0] == null) {
                throw new ArgumentException("Data rows cannot be null.", nameof(items));
            }

            for (var index = 1; index < items.Count; index++) {
                if (items[index] == null) {
                    throw new InvalidOperationException("Data rows cannot contain null entries.");
                }
            }
        }

        private static void ValidateFirstRowShape(object first) {
            if (TabularDataTableBuilder.IsScalarValue(first)) {
                throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries.");
            }

            bool hasColumns = first switch {
                IReadOnlyDictionary<string, object?> readOnlyDictionary => readOnlyDictionary.Count > 0,
                IDictionary<string, object?> genericDictionary => genericDictionary.Count > 0,
                IDictionary dictionary => dictionary.Count > 0,
                _ => first.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Any(property => property.CanRead && property.GetIndexParameters().Length == 0)
            };

            if (!hasColumns) {
                throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries.");
            }
        }

        private static object? NormalizeCellValue(object? value, ObjectDataTableBuilderOptions options) {
            if (value == null || value == DBNull.Value || !options.NormalizeCollectionValues) {
                return value;
            }

            if (value is string || value is IDictionary) {
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
    }
}
