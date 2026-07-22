#nullable enable

using System.Data;
using System.Diagnostics.CodeAnalysis;

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

        if (_mode == CsvLoadMode.Stream && _streamingSource is not null)
        {
            if (options.Schema is not null || _schema is not null || options.InferSchema)
            {
                return ToStreamingDataTable(options);
            }

            return ToStreamingStringDataTable(options.TableName);
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

    private DataTable ToStreamingStringDataTable(string? tableName)
    {
        var table = CreateDataTable(tableName, schemaColumns: null);
        var values = new object?[_header.Count];
        table.BeginLoadData();
        try
        {
            foreach (var row in _streamingSource!.ReadReusableStringRows())
            {
                FillStreamingStringDataTableRow(row, values);
                table.Rows.Add(values);
            }
        }
        finally
        {
            table.EndLoadData();
        }

        return table;
    }

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "CSV schema projection creates scalar DataColumns only and never uses DataColumn expressions; DataTable.Load therefore has no expression-member preservation requirement here.")]
    private DataTable ToStreamingDataTable(CsvDataTableOptions options)
    {
        using var reader = CreateDataReader(new CsvDataReaderOptions
        {
            Schema = options.Schema,
            InferSchema = options.InferSchema,
            SchemaSampleSize = options.SchemaSampleSize
        });
        var table = new DataTable(ResolveDataTableName(options.TableName));
        table.Load(reader);
        return table;
    }

    private void FillStreamingStringDataTableRow(IReadOnlyList<string> row, object?[] values)
    {
        var options = _streamingSource!.Options;
        var sourceColumnCount = _streamingSource.SourceColumnCount;
        if (options.ColumnCountMismatchPolicy == CsvColumnCountMismatchPolicy.Strict &&
            row.Count != sourceColumnCount)
        {
            throw new CsvException($"Row contains {row.Count} values but header defines {sourceColumnCount} columns.");
        }

        var copyCount = Math.Min(row.Count, sourceColumnCount);
        for (var i = 0; i < copyCount; i++)
        {
            var value = row[i];
            values[i] = options.NullValue is not null && string.Equals(value, options.NullValue, StringComparison.Ordinal)
                ? DBNull.Value
                : value;
        }

        for (var i = copyCount; i < sourceColumnCount; i++)
        {
            values[i] = string.Empty;
        }

        if (options.StaticColumns is null || options.StaticColumns.Count == 0)
        {
            return;
        }

        var index = sourceColumnCount;
        foreach (var staticColumn in options.StaticColumns)
        {
            values[index++] = staticColumn.Value is null || staticColumn.Value == DBNull.Value
                ? DBNull.Value
                : Convert.ToString(staticColumn.Value, options.Culture) ?? string.Empty;
        }
    }

    [UnconditionalSuppressMessage("Trimming", "IL2072", Justification = "CSV schema types are normalized to the supported scalar framework type set before DataColumn creation.")]
    private DataTable CreateDataTable(string? tableName, IReadOnlyDictionary<string, CsvSchemaColumn>? schemaColumns)
    {
        var table = new DataTable(ResolveDataTableName(tableName));
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

    private static string ResolveDataTableName(string? tableName) =>
        string.IsNullOrWhiteSpace(tableName) ? "CsvData" : tableName!;
}
