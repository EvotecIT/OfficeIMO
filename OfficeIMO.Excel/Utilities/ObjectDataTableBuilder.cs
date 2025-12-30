using System.Data;
using System.Diagnostics.CodeAnalysis;
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

            var list = items.ToList();
            if (list.Count == 0) {
                throw new ArgumentException("Provide at least one data row.", nameof(items));
            }

            var first = list.FirstOrDefault();
            if (first == null) {
                throw new ArgumentException("Data rows cannot be null.", nameof(items));
            }

            var columns = ObjectDataHelpers.GetColumnNames(first);
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
                    row[column] = ObjectDataHelpers.GetValue(item, column) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }

            return table;
        }

    }
}
