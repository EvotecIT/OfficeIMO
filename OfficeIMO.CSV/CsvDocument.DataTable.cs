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
        var columns = CsvDataProjectionBuilder.Create(_header, schemaColumns);
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
                    values[i] = CsvDataProjectionConverter.ConvertValue(value, column, rowIndex, _culture, _dateTimeFormats);
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

    private DataTable CreateDataTable(string? tableName, IReadOnlyDictionary<string, CsvSchemaColumn>? schemaColumns)
    {
        var table = new DataTable(string.IsNullOrWhiteSpace(tableName) ? "CsvData" : tableName);
        foreach (var columnName in _header)
        {
            CsvSchemaColumn? schemaColumn = null;
            schemaColumns?.TryGetValue(columnName, out schemaColumn);
            var dataType = CsvDataProjectionBuilder.ResolveDataColumnType(schemaColumn?.DataType);
            var column = table.Columns.Add(columnName, dataType);
            if (schemaColumn is not null)
            {
                column.AllowDBNull = !schemaColumn.IsRequired;
            }
        }

        return table;
    }
}
