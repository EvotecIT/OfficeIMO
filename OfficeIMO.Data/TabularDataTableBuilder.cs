using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

namespace OfficeIMO.Data;

/// <summary>
/// Converts common object, dictionary, DataTable, DataView, IDataReader, DataRow, and IDataRecord inputs into tabular data.
/// </summary>
public static class TabularDataTableBuilder {
    /// <summary>Converts an input sequence into a DataTable.</summary>
    public static DataTable FromItems(IEnumerable<object?> input, TabularDataOptions? options = null) {
        if (input == null) {
            throw new ArgumentNullException(nameof(input));
        }

        options ??= new TabularDataOptions();
        var items = input
            .Select(item => Unwrap(item, options))
            .Where(item => options.PreserveNullRows || item != null)
            .ToList();

        if (items.Count == 1 && options.ExpandSingleEnumerableInput && ShouldExpandSingleEnumerableInput(items[0])) {
            items = ((IEnumerable)items[0]!)
                .Cast<object?>()
                .Select(item => Unwrap(item, options))
                .Where(item => options.PreserveNullRows || item != null)
                .ToList();
        }

        if (items.Count == 0) {
            return CreateTable(options.TableName);
        }

        if (items.Count == 1) {
            var single = items[0];
            if (single is DataTable dataTable) {
                return options.CopyExistingDataTable ? dataTable.Copy() : dataTable;
            }

            if (single is DataView dataView) {
                var table = dataView.ToTable();
                ApplyTableName(table, options.TableName);
                return table;
            }

            if (single is IDataReader reader) {
                var table = CreateTable(options.TableName);
                table.Load(reader);
                return table;
            }
        }

        if (items[0] is DataRow firstRow) {
            return FromDataRows(items, firstRow.Table, options.TableName);
        }

        if (items[0] is DataRowView firstRowView) {
            return FromDataRowViews(items, firstRowView.Row.Table, options.TableName);
        }

        if (items[0] is IDataRecord firstRecord) {
            return FromDataRecords(items, firstRecord, options.TableName);
        }

        return FromObjects(items, options);
    }

    /// <summary>Returns the only IDataReader from an input sequence, or null when the input is not exactly one reader.</summary>
    public static IDataReader? TryGetSingleDataReader(IEnumerable<object?> input, Func<object?, object?>? unwrapValue = null) {
        if (input == null) {
            throw new ArgumentNullException(nameof(input));
        }

        IDataReader? reader = null;
        var count = 0;
        foreach (var item in input) {
            if (item == null) {
                continue;
            }

            count++;
            if (count > 1) {
                return null;
            }

            reader = (unwrapValue?.Invoke(item) ?? item) as IDataReader;
        }

        return count == 1 ? reader : null;
    }

    /// <summary>Returns the only DataSet from an input sequence, or null when the input is not exactly one DataSet.</summary>
    public static DataSet? TryGetSingleDataSet(IEnumerable<object?> input, Func<object?, object?>? unwrapValue = null) {
        if (input == null) {
            throw new ArgumentNullException(nameof(input));
        }

        DataSet? dataSet = null;
        var count = 0;
        foreach (var item in input) {
            if (item == null) {
                continue;
            }

            count++;
            if (count > 1) {
                return null;
            }

            dataSet = (unwrapValue?.Invoke(item) ?? item) as DataSet;
        }

        return count == 1 ? dataSet : null;
    }

    /// <summary>Returns true when the value is handled as a scalar row value.</summary>
    public static bool IsScalarValue(object? item) {
        if (item == null || item == DBNull.Value) {
            return true;
        }

        var type = item.GetType();
        return type.IsPrimitive ||
               type.IsEnum ||
               item is string or decimal or DateTime or DateTimeOffset or TimeSpan or Guid;
    }

    private static object? Unwrap(object? item, TabularDataOptions options)
        => options.UnwrapValue?.Invoke(item) ?? item;

    private static DataTable FromDataRows(IReadOnlyList<object?> items, DataTable source, string? tableName) {
        var table = CloneTable(source, tableName);
        foreach (var item in items) {
            if (item is not DataRow row) {
                throw new ArgumentException("DataRow input cannot be mixed with other input types.", nameof(items));
            }

            AddCompatibleDataRow(table, row);
        }

        return table;
    }

    private static DataTable FromDataRowViews(IReadOnlyList<object?> items, DataTable source, string? tableName) {
        var table = CloneTable(source, tableName);
        foreach (var item in items) {
            if (item is not DataRowView rowView) {
                throw new ArgumentException("DataRowView input cannot be mixed with other input types.", nameof(items));
            }

            AddCompatibleDataRow(table, rowView.Row);
        }

        return table;
    }

    private static DataTable FromDataRecords(IReadOnlyList<object?> items, IDataRecord firstRecord, string? tableName) {
        var table = CreateTable(tableName);
        var columns = GetDataRecordColumns(firstRecord);
        foreach (var column in columns) {
            table.Columns.Add(column.TableName, column.FieldType);
        }

        foreach (var item in items) {
            if (item is not IDataRecord record) {
                throw new ArgumentException("IDataRecord input cannot be mixed with other input types.", nameof(items));
            }

            EnsureCompatibleDataRecord(record, columns);
            var row = table.NewRow();
            for (var index = 0; index < columns.Count; index++) {
                row[index] = record.IsDBNull(index) ? DBNull.Value : record.GetValue(index);
            }

            table.Rows.Add(row);
        }

        return table;
    }

    private static IReadOnlyList<DataRecordColumn> GetDataRecordColumns(IDataRecord record) {
        var columns = new List<DataRecordColumn>(record.FieldCount);
        var tableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < record.FieldCount; index++) {
            var sourceName = record.GetName(index);
            var baseName = string.IsNullOrWhiteSpace(sourceName)
                ? $"Column{index + 1}"
                : sourceName;
            var tableName = GetUniqueColumnName(baseName, tableNames);
            columns.Add(new DataRecordColumn(sourceName, tableName, record.GetFieldType(index) ?? typeof(object)));
        }

        return columns;
    }

    private static void EnsureCompatibleDataRecord(IDataRecord record, IReadOnlyList<DataRecordColumn> expectedColumns) {
        if (record.FieldCount != expectedColumns.Count) {
            throw new ArgumentException("IDataRecord inputs must have the same field count.", nameof(record));
        }

        for (var index = 0; index < expectedColumns.Count; index++) {
            var expected = expectedColumns[index];
            if (!string.Equals(record.GetName(index), expected.SourceName, StringComparison.OrdinalIgnoreCase) ||
                (record.GetFieldType(index) ?? typeof(object)) != expected.FieldType) {
                throw new ArgumentException("IDataRecord inputs must have compatible column schemas.", nameof(record));
            }
        }
    }

    private static DataTable FromObjects(IReadOnlyList<object?> items, TabularDataOptions options) {
        var rows = new List<IReadOnlyDictionary<string, object?>>(items.Count);
        var columns = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var item in items) {
            var row = GetProperties(item, options);
            rows.Add(row);

            if (options.ColumnDiscoveryMode == TabularColumnDiscoveryMode.AllRows || rows.Count == 1) {
                foreach (var key in row.Keys.Where(seen.Add)) {
                    columns.Add(key);
                }
            }
        }

        if (columns.Count == 0) {
            columns.Add(GetScalarColumnName(options));
            rows.Clear();
            foreach (var item in items) {
                rows.Add(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
                    [GetScalarColumnName(options)] = NormalizeValue(item, options)
                });
            }
        }

        var table = CreateTable(options.TableName);
        foreach (var column in columns) {
            table.Columns.Add(column, InferColumnType(rows, column));
        }

        table.MinimumCapacity = Math.Max(table.MinimumCapacity, rows.Count);
        table.BeginLoadData();
        try {
            foreach (var rowValues in rows) {
                var row = table.NewRow();
                foreach (var column in columns) {
                    row[column] = rowValues.TryGetValue(column, out var value) && value != null
                        ? value
                        : DBNull.Value;
                }

                table.Rows.Add(row);
            }
        } finally {
            table.EndLoadData();
        }

        return table;
    }

    private static IReadOnlyDictionary<string, object?> GetProperties(object? item, TabularDataOptions options) {
        var projected = options.ProjectObject?.Invoke(item);
        if (projected != null) {
            return NormalizeDictionary(projected, options);
        }

        if (IsScalarValue(item)) {
            return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
                [GetScalarColumnName(options)] = NormalizeValue(item, options)
            };
        }

        if (item is IDictionary dictionary) {
            var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (DictionaryEntry entry in dictionary) {
                if (entry.Key != null) {
                    values[entry.Key.ToString()!] = NormalizeValue(entry.Value, options);
                }
            }

            return values;
        }

        return ProjectPublicProperties(item!, options);
    }

    private static IReadOnlyDictionary<string, object?> NormalizeDictionary(IReadOnlyDictionary<string, object?> source, TabularDataOptions options) {
        var result = new Dictionary<string, object?>(source.Count, StringComparer.OrdinalIgnoreCase);
        foreach (var entry in source) {
            if (!string.IsNullOrWhiteSpace(entry.Key)) {
                result[entry.Key] = NormalizeValue(entry.Value, options);
            }
        }

        return result;
    }

    private static IReadOnlyDictionary<string, object?> ProjectPublicProperties(object item, TabularDataOptions options) {
        var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in item.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)) {
            if (!property.CanRead || property.GetIndexParameters().Length != 0) {
                continue;
            }

            result[property.Name] = NormalizeValue(property.GetValue(item), options);
        }

        return result;
    }

    private static Type InferColumnType(IReadOnlyList<IReadOnlyDictionary<string, object?>> rows, string column) {
        Type? inferred = null;
        foreach (var row in rows) {
            if (!row.TryGetValue(column, out var value) || value == null || value == DBNull.Value) {
                continue;
            }

            var valueType = value.GetType();
            if (inferred == null) {
                inferred = valueType;
                continue;
            }

            if (inferred != valueType) {
                return typeof(object);
            }
        }

        return inferred ?? typeof(object);
    }

    private static object? NormalizeValue(object? value, TabularDataOptions options)
        => options.NormalizeValue?.Invoke(value) ?? value;

    private static bool ShouldExpandSingleEnumerableInput(object? item)
        => item is IEnumerable &&
           item is not string &&
           item is not byte[] &&
           item is not DataSet &&
           item is not DataTable &&
           item is not DataView &&
           item is not IDataReader &&
           item is not IDataRecord &&
           item is not IDictionary;

    private static DataTable CloneTable(DataTable source, string? tableName) {
        var table = source.Clone();
        ApplyTableName(table, tableName);
        return table;
    }

    private static void AddCompatibleDataRow(DataTable table, DataRow row) {
        EnsureCompatibleSchema(table, row.Table);

        var newRow = table.NewRow();
        foreach (DataColumn column in table.Columns) {
            newRow[column.ColumnName] = row[column.ColumnName];
        }

        table.Rows.Add(newRow);
    }

    private static void EnsureCompatibleSchema(DataTable target, DataTable source) {
        if (source.Columns.Count != target.Columns.Count) {
            throw new ArgumentException("DataRow inputs must have compatible column schemas.", nameof(source));
        }

        foreach (DataColumn targetColumn in target.Columns) {
            if (!source.Columns.Contains(targetColumn.ColumnName)) {
                throw new ArgumentException("DataRow inputs must have compatible column schemas.", nameof(source));
            }

            var sourceColumn = source.Columns[targetColumn.ColumnName]!;
            if (sourceColumn.DataType != targetColumn.DataType) {
                throw new ArgumentException("DataRow inputs must have compatible column schemas.", nameof(source));
            }
        }
    }

    private static DataTable CreateTable(string? tableName)
        => string.IsNullOrWhiteSpace(tableName) ? new DataTable() : new DataTable(tableName);

    private static void ApplyTableName(DataTable table, string? tableName) {
        if (!string.IsNullOrWhiteSpace(tableName)) {
            table.TableName = tableName;
        }
    }

    private static string GetUniqueColumnName(string columnName, HashSet<string> seen) {
        var candidate = columnName;
        var suffix = 2;
        while (!seen.Add(candidate)) {
            candidate = $"{columnName}_{suffix}";
            suffix++;
        }

        return candidate;
    }

    private static string GetScalarColumnName(TabularDataOptions options)
        => string.IsNullOrWhiteSpace(options.ScalarColumnName) ? "Value" : options.ScalarColumnName;

    private sealed class DataRecordColumn {
        internal DataRecordColumn(string sourceName, string tableName, Type fieldType) {
            SourceName = sourceName;
            TableName = tableName;
            FieldType = fieldType;
        }

        internal string SourceName { get; }

        internal string TableName { get; }

        internal Type FieldType { get; }
    }
}
