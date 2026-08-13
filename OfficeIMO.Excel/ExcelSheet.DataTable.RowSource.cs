using System.Data;
using System.IO;

namespace OfficeIMO.Excel {
    internal sealed class DataTableTabularRowSource : IExcelSheetTabularRowSource {
        private readonly DataTable _table;
        private readonly DataColumn?[] _columns;
        private readonly string[] _headers;

        internal DataTableTabularRowSource(
            DataTable table,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<string> headers) {
            _table = table ?? throw new ArgumentNullException(nameof(table));
            if (columnNames == null) throw new ArgumentNullException(nameof(columnNames));
            if (headers == null) throw new ArgumentNullException(nameof(headers));
            if (columnNames.Count != headers.Count) {
                throw new ArgumentException("Column and header counts must match.", nameof(headers));
            }

            var availableColumns = table.Columns
                .Cast<DataColumn>()
                .ToDictionary(column => column.ColumnName, StringComparer.Ordinal);
            _columns = new DataColumn?[columnNames.Count];
            _headers = new string[headers.Count];
            for (int index = 0; index < columnNames.Count; index++) {
                string name = columnNames[index];
                if (!availableColumns.TryGetValue(name, out DataColumn? column)) {
                    column = FindCaseInsensitiveColumn(table.Columns, name);
                }
                _columns[index] = column;
                _headers[index] = headers[index];
            }
        }

        private static DataColumn? FindCaseInsensitiveColumn(DataColumnCollection columns, string name) {
            DataColumn? match = null;
            foreach (DataColumn column in columns) {
                if (!string.Equals(column.ColumnName, name, StringComparison.OrdinalIgnoreCase)) continue;
                if (match != null) {
                    throw new InvalidDataException(
                        $"DataTable column '{name}' is ambiguous because the schema contains case-distinct matches. Use the exact column casing in the projection.");
                }
                match = column;
            }
            return match;
        }

        public int ColumnCount => _columns.Length;

        public int RowCount => _table.Rows.Count;

        public string GetColumnName(int index) => _headers[index];

        public Type GetColumnType(int index) => _columns[index]?.DataType ?? typeof(object);

        public object? GetValue(int rowIndex, int columnIndex) {
            DataRow row = _table.Rows[rowIndex];
            DataColumn? column = _columns[columnIndex];
            if (column == null) return null;
            return row.IsNull(column) ? null : row[column];
        }

        public bool TryGetBufferedRow(int rowIndex, out object?[]? values) {
            values = null;
            return false;
        }

        public bool TryGetFlatValues(out object?[] values, out int columnCount) {
            values = Array.Empty<object?>();
            columnCount = 0;
            return false;
        }
    }
}
