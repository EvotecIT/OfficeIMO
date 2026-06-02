using System.Data;
using System.Globalization;
using System.ComponentModel;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private sealed partial class DirectDataSetTableModel : IExcelSheetTabularRowSource {
            private const int MaxAutoFitStringWidthCacheEntriesPerColumn = 1024;
            private const long BufferedDictionaryCellLimit = 500_000;

            private enum AutoFitWidthKind {
                Object,
                String,
                Boolean,
                DateTime,
                DateTimeOffset,
                TimeSpan,
                Double,
                Float,
                Decimal,
                SByte,
                Byte,
                Int16,
                UInt16,
                Int32,
                UInt32,
                Int64,
                UInt64,
#if NET6_0_OR_GREATER
                DateOnly,
                TimeOnly,
#endif
            }

            private readonly DataTable? _sourceTable;
            private readonly DirectDataSetColumnModel[]? _columns;
            private readonly int[]? _stringCandidateColumnIndexes;
            private readonly object?[][]? _arrayRows;
            private readonly List<object?[]>? _listRows;
            private readonly DirectCellValueRows _cellValueRows;
            private readonly bool _hasCellValueRows;
            private readonly IReadOnlyList<Dictionary<string, object?>>? _exactDictionaryRows;
            private readonly IReadOnlyList<IReadOnlyDictionary<string, object?>>? _dictionaryRows;
            private readonly IReadOnlyList<System.Collections.IDictionary>? _legacyDictionaryRows;
            private readonly bool _legacyDictionaryExactKeyLookup;
            private string[]? _columnNameArray;

            private DirectDataSetTableModel(DataTable sourceTable) {
                _sourceTable = sourceTable;
                _columns = CreateColumns(sourceTable);
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(_columns);
            }

            private DirectDataSetTableModel(DataTable sourceTable, DirectDataSetColumnModel[] columns) {
                _sourceTable = sourceTable;
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
            }

            private DirectDataSetTableModel(DirectDataSetColumnModel[] columns, IReadOnlyList<object?[]> rows) {
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
                if (rows is object?[][] arrayRows) {
                    _arrayRows = arrayRows;
                } else if (rows is List<object?[]> listRows) {
                    _listRows = listRows;
                } else {
                    var copiedRows = new object?[rows.Count][];
                    for (int i = 0; i < copiedRows.Length; i++) {
                        copiedRows[i] = rows[i];
                    }

                    _arrayRows = copiedRows;
                }
            }

            private DirectDataSetTableModel(DirectDataSetColumnModel[] columns, DirectCellValueRows cellValueRows) {
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
                _cellValueRows = cellValueRows;
                _hasCellValueRows = true;
            }

            private DirectDataSetTableModel(DirectDataSetColumnModel[] columns, IReadOnlyList<IReadOnlyDictionary<string, object?>> dictionaryRows) {
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
                _dictionaryRows = dictionaryRows;
            }

            private DirectDataSetTableModel(DirectDataSetColumnModel[] columns, IReadOnlyList<Dictionary<string, object?>> exactDictionaryRows) {
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
                _exactDictionaryRows = exactDictionaryRows;
            }

            private DirectDataSetTableModel(DirectDataSetColumnModel[] columns, IReadOnlyList<System.Collections.IDictionary> legacyDictionaryRows, bool exactKeyLookup = false) {
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
                _legacyDictionaryRows = legacyDictionaryRows;
                _legacyDictionaryExactKeyLookup = exactKeyLookup;
            }

            internal static DirectDataSetTableModel Reference(DataTable table) => new DirectDataSetTableModel(table);

            internal static DirectDataSetTableModel FromRows(IReadOnlyList<string> columnNames, IReadOnlyList<Type> columnTypes, IReadOnlyList<object?[]> rows) {
                if (columnNames.Count != columnTypes.Count) {
                    throw new ArgumentException("Column name and type counts must match.", nameof(columnTypes));
                }

                var columns = new DirectDataSetColumnModel[columnNames.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(columnNames[i], columnTypes[i]);
                }

                return new DirectDataSetTableModel(columns, rows);
            }

            internal static DirectDataSetTableModel FromCellValues(
                IReadOnlyList<string> columnNames,
                IReadOnlyList<Type> columnTypes,
                object?[] values,
                int columnCount,
                int rowCount,
                bool valuesMatchColumnTypes) {
                if (columnNames.Count != columnTypes.Count) {
                    throw new ArgumentException("Column name and type counts must match.", nameof(columnTypes));
                }

                if (columnNames.Count != columnCount) {
                    throw new ArgumentException("Column count must match the column metadata.", nameof(columnCount));
                }

                var columns = new DirectDataSetColumnModel[columnNames.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(columnNames[i], columnTypes[i]);
                }

                return new DirectDataSetTableModel(columns, new DirectCellValueRows(values, columnCount, rowCount, valuesMatchColumnTypes));
            }

            internal static DirectDataSetTableModel FromLegacyDictionaries(IReadOnlyList<string> columnNames, IReadOnlyList<Type> columnTypes, IReadOnlyList<System.Collections.IDictionary> rows) {
                if (columnNames.Count != columnTypes.Count) {
                    throw new ArgumentException("Column name and type counts must match.", nameof(columnTypes));
                }

                var columns = new DirectDataSetColumnModel[columnNames.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(columnNames[i], columnTypes[i]);
                }

                return new DirectDataSetTableModel(columns, rows, HasCaseInsensitiveDuplicateColumnNames(columnNames));
            }

            internal static DirectDataSetTableModel FromExactDictionaries(IReadOnlyList<string> columnNames, IReadOnlyList<Type> columnTypes, IReadOnlyList<Dictionary<string, object?>> rows) {
                if (columnNames.Count != columnTypes.Count) {
                    throw new ArgumentException("Column name and type counts must match.", nameof(columnTypes));
                }

                var columns = new DirectDataSetColumnModel[columnNames.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(columnNames[i], columnTypes[i]);
                }

                if (ShouldBufferDictionaryRows(rows.Count, columns.Length)) {
                    return new DirectDataSetTableModel(columns, SnapshotExactDictionaryRows(columnNames, rows));
                }

                return new DirectDataSetTableModel(columns, rows);
            }

            internal static DirectDataSetTableModel FromDictionaries(IReadOnlyList<string> columnNames, IReadOnlyList<Type> columnTypes, IReadOnlyList<IReadOnlyDictionary<string, object?>> rows) {
                if (columnNames.Count != columnTypes.Count) {
                    throw new ArgumentException("Column name and type counts must match.", nameof(columnTypes));
                }

                var columns = new DirectDataSetColumnModel[columnNames.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(columnNames[i], columnTypes[i]);
                }

                if (ShouldBufferDictionaryRows(rows.Count, columns.Length)) {
                    return new DirectDataSetTableModel(columns, SnapshotDictionaryRows(columnNames, rows));
                }

                return new DirectDataSetTableModel(columns, rows);
            }

            private static bool ShouldBufferDictionaryRows(int rowCount, int columnCount)
                => rowCount > 0
                    && columnCount > 0
                    && (long)rowCount * columnCount <= BufferedDictionaryCellLimit;

            private static object?[][] SnapshotExactDictionaryRows(
                IReadOnlyList<string> columnNames,
                IReadOnlyList<Dictionary<string, object?>> rows) {
                var bufferedRows = new object?[rows.Count][];
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                    Dictionary<string, object?> source = rows[rowIndex];
                    var values = new object?[columnNames.Count];
                    for (int columnIndex = 0; columnIndex < columnNames.Count; columnIndex++) {
                        values[columnIndex] = source.TryGetValue(columnNames[columnIndex], out object? value)
                            ? value
                            : null;
                    }

                    bufferedRows[rowIndex] = values;
                }

                return bufferedRows;
            }

            private static object?[][] SnapshotDictionaryRows(
                IReadOnlyList<string> columnNames,
                IReadOnlyList<IReadOnlyDictionary<string, object?>> rows) {
                var bufferedRows = new object?[rows.Count][];
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                    IReadOnlyDictionary<string, object?> source = rows[rowIndex];
                    var values = new object?[columnNames.Count];
                    for (int columnIndex = 0; columnIndex < columnNames.Count; columnIndex++) {
                        values[columnIndex] = source.TryGetValue(columnNames[columnIndex], out object? value)
                            ? value
                            : null;
                    }

                    bufferedRows[rowIndex] = values;
                }

                return bufferedRows;
            }

            internal DirectDataSetTableModel WithGeneratedColumnNames() {
                var columns = new DirectDataSetColumnModel[ColumnCount];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel("Column" + (i + 1).ToString(CultureInfo.InvariantCulture), GetColumnType(i));
                }

                if (_sourceTable != null) {
                    return new DirectDataSetTableModel(_sourceTable, columns);
                }

                if (_exactDictionaryRows != null) {
                    return new DirectDataSetTableModel(columns, _exactDictionaryRows);
                }

                if (_dictionaryRows != null) {
                    return new DirectDataSetTableModel(columns, _dictionaryRows);
                }

                if (_legacyDictionaryRows != null) {
                    return new DirectDataSetTableModel(columns, _legacyDictionaryRows, _legacyDictionaryExactKeyLookup);
                }

                if (_hasCellValueRows) {
                    return new DirectDataSetTableModel(columns, _cellValueRows);
                }

                return new DirectDataSetTableModel(columns, GetBufferedRowsForReuse());
            }

            internal static DirectDataSetTableModel Snapshot(DataTable table, CancellationToken ct) {
                var columns = CreateColumns(table);
                if (columns.Length == 8) {
                    return new DirectDataSetTableModel(columns, SnapshotEightColumnRows(table, ct));
                }

                var rows = new object?[table.Rows.Count][];
                bool canCancel = ct.CanBeCanceled;
                for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    DataRow row = table.Rows[rowIndex];
                    var values = new object?[columns.Length];
                    for (int columnIndex = 0; columnIndex < columns.Length; columnIndex++) {
                        object? value = row[columnIndex];
                        values[columnIndex] = value == DBNull.Value ? null : value;
                    }

                    rows[rowIndex] = values;
                }

                return new DirectDataSetTableModel(columns, rows);
            }

            private static object?[][] SnapshotEightColumnRows(DataTable table, CancellationToken ct) {
                var rows = new object?[table.Rows.Count][];
                bool canCancel = ct.CanBeCanceled;
                for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    DataRow row = table.Rows[rowIndex];
                    object? value0 = row[0];
                    object? value1 = row[1];
                    object? value2 = row[2];
                    object? value3 = row[3];
                    object? value4 = row[4];
                    object? value5 = row[5];
                    object? value6 = row[6];
                    object? value7 = row[7];
                    rows[rowIndex] = new object?[] {
                        value0 == DBNull.Value ? null : value0,
                        value1 == DBNull.Value ? null : value1,
                        value2 == DBNull.Value ? null : value2,
                        value3 == DBNull.Value ? null : value3,
                        value4 == DBNull.Value ? null : value4,
                        value5 == DBNull.Value ? null : value5,
                        value6 == DBNull.Value ? null : value6,
                        value7 == DBNull.Value ? null : value7
                    };
                }

                return rows;
            }

            private static DirectDataSetColumnModel[] CreateColumns(DataTable table) {
                var columns = new DirectDataSetColumnModel[table.Columns.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(table.Columns[i].ColumnName, table.Columns[i].DataType);
                }

                return columns;
            }

            internal int ColumnCount => _columns!.Length;

            internal int RowCount => _sourceTable?.Rows.Count ?? _arrayRows?.Length ?? _listRows?.Count ?? (_hasCellValueRows ? _cellValueRows.Count : (int?)null) ?? _exactDictionaryRows?.Count ?? _dictionaryRows?.Count ?? _legacyDictionaryRows!.Count;

            internal string GetColumnName(int index) => _columns![index].Name;

            internal Type GetColumnType(int index) => _columns![index].DataType;

            int IExcelSheetTabularRowSource.ColumnCount => ColumnCount;

            int IExcelSheetTabularRowSource.RowCount => RowCount;

            string IExcelSheetTabularRowSource.GetColumnName(int index) => GetColumnName(index);

            Type IExcelSheetTabularRowSource.GetColumnType(int index) => GetColumnType(index);

            object? IExcelSheetTabularRowSource.GetValue(int rowIndex, int columnIndex) => GetValue(rowIndex, columnIndex);

            bool IExcelSheetTabularRowSource.TryGetBufferedRow(int rowIndex, out object?[]? values) {
                values = GetBufferedRow(rowIndex);
                return values != null;
            }

            bool IExcelSheetTabularRowSource.TryGetFlatValues(out object?[] values, out int columnCount) {
                if (_hasCellValueRows) {
                    values = _cellValueRows.Values;
                    columnCount = _cellValueRows.ColumnCount;
                    return true;
                }

                values = Array.Empty<object?>();
                columnCount = 0;
                return false;
            }

            internal string[] CreateColumnNameArray() {
                if (_columnNameArray != null) {
                    return _columnNameArray;
                }

                var columnNames = new string[_columns!.Length];
                for (int i = 0; i < columnNames.Length; i++) {
                    columnNames[i] = _columns[i].Name;
                }

                _columnNameArray = columnNames;
                return columnNames;
            }

            private static bool HasCaseInsensitiveDuplicateColumnNames(IReadOnlyList<string> columnNames) {
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < columnNames.Count; i++) {
                    if (!seen.Add(columnNames[i] ?? string.Empty)) {
                        return true;
                    }
                }

                return false;
            }

            internal int[]? GetStringCandidateColumnIndexes() => _stringCandidateColumnIndexes;

            private static int[]? CreateStringCandidateColumnIndexes(DirectDataSetColumnModel[] columns) {
                int[]? indexes = null;
                int count = 0;
                for (int i = 0; i < columns.Length; i++) {
                    Type dataType = columns[i].DataType;
                    if (dataType != typeof(string) && dataType != typeof(object)) {
                        continue;
                    }

                    indexes ??= new int[columns.Length];
                    indexes[count++] = i;
                }

                if (indexes == null) {
                    return null;
                }

                if (count == indexes.Length) {
                    return indexes;
                }

                Array.Resize(ref indexes, count);
                return indexes;
            }

            internal DataRow? GetSourceRow(int rowIndex) => _sourceTable?.Rows[rowIndex];

            internal bool HasSourceRows => _sourceTable != null;

            internal bool TryGetExactDictionaryRows(out IReadOnlyList<Dictionary<string, object?>> rows) {
                if (_exactDictionaryRows != null) {
                    rows = _exactDictionaryRows;
                    return true;
                }

                rows = Array.Empty<Dictionary<string, object?>>();
                return false;
            }

            internal bool TryGetDictionaryRows(out IReadOnlyList<IReadOnlyDictionary<string, object?>> rows) {
                if (_dictionaryRows != null) {
                    rows = _dictionaryRows;
                    return true;
                }

                rows = Array.Empty<IReadOnlyDictionary<string, object?>>();
                return false;
            }

            internal bool TryGetLegacyDictionaryRows(out IReadOnlyList<System.Collections.IDictionary> rows) {
                if (_legacyDictionaryRows != null) {
                    rows = _legacyDictionaryRows;
                    return true;
                }

                rows = Array.Empty<System.Collections.IDictionary>();
                return false;
            }

            internal bool TryGetBufferedRows(out DirectBufferedRows rows) {
                if (_arrayRows != null) {
                    rows = new DirectBufferedRows(_arrayRows);
                    return true;
                }

                if (_listRows != null) {
                    rows = new DirectBufferedRows(_listRows);
                    return true;
                }

                rows = default;
                return false;
            }

            internal bool TryGetCellValueRows(out DirectCellValueRows rows) {
                if (_hasCellValueRows) {
                    rows = _cellValueRows;
                    return true;
                }

                rows = default;
                return false;
            }

            internal object?[]? GetBufferedRow(int rowIndex) {
                if (_arrayRows != null) {
                    return _arrayRows[rowIndex];
                }

                return _listRows?[rowIndex];
            }

            internal object? GetValue(int rowIndex, int columnIndex) {
                object? value;
                if (_sourceTable != null) {
                    value = _sourceTable.Rows[rowIndex][columnIndex];
                } else if (_exactDictionaryRows != null) {
                    value = _exactDictionaryRows[rowIndex].TryGetValue(GetColumnName(columnIndex), out object? dictionaryValue)
                        ? dictionaryValue
                        : null;
                } else if (_dictionaryRows != null) {
                    value = _dictionaryRows[rowIndex].TryGetValue(GetColumnName(columnIndex), out object? dictionaryValue)
                        ? dictionaryValue
                        : null;
                } else if (_legacyDictionaryRows != null) {
                    value = GetLegacyDictionaryValue(_legacyDictionaryRows[rowIndex], GetColumnName(columnIndex), _legacyDictionaryExactKeyLookup);
                } else if (_hasCellValueRows) {
                    value = _cellValueRows.GetValue(rowIndex, columnIndex);
                } else {
                    value = GetBufferedRow(rowIndex)![columnIndex];
                }

                return value == DBNull.Value ? null : value;
            }

            internal static object? GetLegacyDictionaryValue(System.Collections.IDictionary dictionary, string column, bool exactKeyLookup = false) {
                if (dictionary.Contains(column)) {
                    return dictionary[column];
                }

                if (exactKeyLookup) {
                    return null;
                }

                foreach (System.Collections.DictionaryEntry entry in dictionary) {
                    string? key = entry.Key?.ToString();
                    if (string.Equals(key, column, StringComparison.OrdinalIgnoreCase)) {
                        return entry.Value;
                    }
                }

                return null;
            }
        }
    }
}
