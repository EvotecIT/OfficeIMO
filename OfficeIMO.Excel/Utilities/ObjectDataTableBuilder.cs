using System.Collections;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

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

            var list = items.ToList();
            if (list.Count == 0) {
                throw new ArgumentException("Provide at least one data row.", nameof(items));
            }

            var first = list.FirstOrDefault();
            if (first == null) {
                throw new ArgumentException("Data rows cannot be null.", nameof(items));
            }

            var columns = GetColumnNames(first);
            if (columns.Count == 0) {
                throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries.");
            }

            var table = new DataTable(tableName);
            foreach (var name in columns) {
                table.Columns.Add(name, typeof(object));
            }

            foreach (var item in list) {
                if (item == null) {
                    throw new InvalidOperationException("Data rows cannot contain null entries.");
                }

                var row = table.NewRow();
                foreach (var column in columns) {
                    row[column] = GetValue(item, column) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }

            return table;
        }

        private static IReadOnlyList<string> GetColumnNames(object item) {
            if (item is IReadOnlyDictionary<string, object?> roDict) {
                return roDict.Keys.Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
            }

            if (item is IDictionary<string, object?> dict) {
                return dict.Keys.Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
            }

            if (item is IDictionary legacyDict) {
                var names = new List<string>();
                foreach (DictionaryEntry entry in legacyDict) {
                    var key = entry.Key?.ToString();
                    if (!string.IsNullOrWhiteSpace(key)) {
                        names.Add(key!);
                    }
                }
                return names;
            }

            var props = item.GetType()
                .GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead && p.GetIndexParameters().Length == 0)
                .OrderBy(p => p.MetadataToken)
                .Select(p => p.Name)
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToList();

            return props;
        }

        private static object? GetValue(object item, string column) {
            if (item is IReadOnlyDictionary<string, object?> roDict) {
                return roDict.TryGetValue(column, out var value) ? value : null;
            }

            if (item is IDictionary<string, object?> dict) {
                return dict.TryGetValue(column, out var value) ? value : null;
            }

            if (item is IDictionary legacyDict) {
                if (legacyDict.Contains(column)) {
                    return legacyDict[column];
                }

                foreach (DictionaryEntry entry in legacyDict) {
                    var key = entry.Key?.ToString();
                    if (string.Equals(key, column, StringComparison.OrdinalIgnoreCase)) {
                        return entry.Value;
                    }
                }

                return null;
            }

            var prop = item.GetType().GetProperty(column, BindingFlags.Public | BindingFlags.Instance);
            return prop?.GetValue(item);
        }
    }
}
