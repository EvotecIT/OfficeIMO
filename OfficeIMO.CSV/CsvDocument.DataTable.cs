#nullable enable

using System.Data;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Projects the CSV document into a <see cref="DataTable"/>.
    /// </summary>
    /// <param name="options">DataTable projection options. When omitted, all columns are emitted as strings.</param>
    /// <returns>A DataTable containing all document rows.</returns>
    public DataTable ToDataTable(CsvDataTableOptions? options = null)
    {
        options ??= new CsvDataTableOptions();
        if (options.SchemaSampleSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "Schema sample size must be greater than zero.");
        }

        var schema = options.Schema ?? _schema ?? (options.InferSchema ? InferSchema(options.SchemaSampleSize) : null);
        var schemaColumns = schema?.Columns.ToDictionary(column => column.Name, StringComparer.OrdinalIgnoreCase);
        var table = CreateDataTable(options.TableName, schemaColumns);
        var columns = CreateDataTableColumnProjections(table, schemaColumns);
        if (_mode == CsvLoadMode.InMemory)
        {
            table.MinimumCapacity = _rows.Count;
        }

        table.BeginLoadData();
        try
        {
            var rowIndex = 0;
            foreach (var row in EnumerateRawRows())
            {
                var values = new object?[columns.Length];
                for (var i = 0; i < columns.Length; i++)
                {
                    var column = columns[i];
                    var value = i < row.Length ? row[i] : null;
                    values[i] = ConvertDataTableValue(value, column.DataType, column.SchemaColumn, rowIndex, column.Name);
                }

                table.Rows.Add(values);
                rowIndex++;
            }
        }
        finally
        {
            table.EndLoadData();
        }

        return table;
    }

    private static DataTableColumnProjection[] CreateDataTableColumnProjections(
        DataTable table,
        IReadOnlyDictionary<string, CsvSchemaColumn>? schemaColumns)
    {
        var columns = new DataTableColumnProjection[table.Columns.Count];
        for (var i = 0; i < columns.Length; i++)
        {
            var dataColumn = table.Columns[i];
            CsvSchemaColumn? schemaColumn = null;
            schemaColumns?.TryGetValue(dataColumn.ColumnName, out schemaColumn);
            columns[i] = new DataTableColumnProjection(dataColumn.ColumnName, dataColumn.DataType, schemaColumn);
        }

        return columns;
    }

    private DataTable CreateDataTable(string? tableName, IReadOnlyDictionary<string, CsvSchemaColumn>? schemaColumns)
    {
        var table = new DataTable(string.IsNullOrWhiteSpace(tableName) ? "CsvData" : tableName);
        foreach (var columnName in _header)
        {
            CsvSchemaColumn? schemaColumn = null;
            schemaColumns?.TryGetValue(columnName, out schemaColumn);
            var dataType = ResolveDataColumnType(schemaColumn?.DataType);
            var column = table.Columns.Add(columnName, dataType);
            if (schemaColumn is not null)
            {
                column.AllowDBNull = !schemaColumn.IsRequired;
            }
        }

        return table;
    }

    private object ConvertDataTableValue(object? value, Type targetType, CsvSchemaColumn? schemaColumn, int rowIndex, string columnName)
    {
        if (IsMissingDataTableValue(value, targetType))
        {
            if (schemaColumn?.DefaultValue is not null)
            {
                value = schemaColumn.DefaultValue;
            }
            else if (schemaColumn?.IsRequired == true)
            {
                throw new CsvException($"Column '{columnName}' is required but row {rowIndex + 1} has no value.");
            }
            else
            {
                return DBNull.Value;
            }
        }

        if (targetType.IsInstanceOfType(value))
        {
            return value!;
        }

        if (!CsvValueConverter.TryConvert(value, targetType, _culture, _dateTimeFormats, out var converted, out var error))
        {
            throw new CsvException($"Column '{columnName}' value on row {rowIndex + 1} cannot be converted to {targetType.Name}: {error}");
        }

        return converted ?? DBNull.Value;
    }

    private readonly struct DataTableColumnProjection
    {
        public DataTableColumnProjection(string name, Type dataType, CsvSchemaColumn? schemaColumn)
        {
            Name = name;
            DataType = dataType;
            SchemaColumn = schemaColumn;
        }

        public string Name { get; }

        public Type DataType { get; }

        public CsvSchemaColumn? SchemaColumn { get; }
    }

    private static Type ResolveDataColumnType(Type? dataType)
    {
        if (dataType is null)
        {
            return typeof(string);
        }

        return Nullable.GetUnderlyingType(dataType) ?? dataType;
    }

    private static bool IsMissingDataTableValue(object? value, Type targetType) =>
        value is null ||
        value == DBNull.Value ||
        (targetType != typeof(string) && value is string { Length: 0 });
}
