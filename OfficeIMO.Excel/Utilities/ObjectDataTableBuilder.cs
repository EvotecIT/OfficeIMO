using System.Collections;
using System.Data;
using System.Diagnostics.CodeAnalysis;
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
            options ??= new ObjectDataTableBuilderOptions();
            return TabularDataTableBuilder.FromItems(items, new TabularDataOptions {
                TableName = tableName,
                ColumnDiscoveryMode = TabularColumnDiscoveryMode.FirstRow,
                NormalizeValue = value => NormalizeCellValue(value, options)
            });
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
