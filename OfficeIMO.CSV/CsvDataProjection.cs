#nullable enable

using System.Globalization;

namespace OfficeIMO.CSV;

internal readonly struct CsvDataColumnProjection
{
    public CsvDataColumnProjection(string name, Type dataType, CsvSchemaColumn? schemaColumn)
    {
        Name = name;
        DataType = dataType;
        SchemaColumn = schemaColumn;
    }

    public string Name { get; }

    public Type DataType { get; }

    public CsvSchemaColumn? SchemaColumn { get; }
}

internal static class CsvDataProjectionConverter
{
    public static object ConvertValue(
        object? value,
        CsvDataColumnProjection column,
        int rowIndex,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats)
    {
        if (IsMissingValue(value, column))
        {
            if (column.SchemaColumn?.DefaultValue is not null)
            {
                value = column.SchemaColumn.DefaultValue;
            }
            else if (column.SchemaColumn?.IsRequired == true)
            {
                throw new CsvException($"Column '{column.Name}' is required but row {rowIndex + 1} has no value.");
            }
            else
            {
                return DBNull.Value;
            }
        }

        if (column.SchemaColumn?.Converter is { } converter)
        {
            value = ConvertWithCustomConverter(value, converter, column.Name, rowIndex, culture);
            if (value is null || value == DBNull.Value)
            {
                return DBNull.Value;
            }
        }

        if (column.DataType.IsInstanceOfType(value))
        {
            return value!;
        }

        if (!CsvValueConverter.TryConvert(value, column.DataType, culture, dateTimeFormats, out var converted, out var error))
        {
            throw new CsvException($"Column '{column.Name}' value on row {rowIndex + 1} cannot be converted to {column.DataType.Name}: {error}");
        }

        return converted ?? DBNull.Value;
    }

    private static object? ConvertWithCustomConverter(
        object? value,
        Func<object?, CultureInfo, object?> converter,
        string columnName,
        int rowIndex,
        CultureInfo culture)
    {
        try
        {
            return converter(value, culture);
        }
        catch (Exception ex) when (ex is not CsvException)
        {
            throw new CsvException($"Column '{columnName}' custom converter failed on row {rowIndex + 1}: {ex.Message}", ex);
        }
    }

    private static bool IsMissingValue(object? value, CsvDataColumnProjection column) =>
        value is null ||
        value == DBNull.Value ||
        (value is string { Length: 0 } &&
            (column.DataType != typeof(string) ||
             column.SchemaColumn?.IsRequired == true ||
             column.SchemaColumn?.DefaultValue is not null));
}

internal static class CsvDataProjectionBuilder
{
    public static CsvDataColumnProjection[] Create(
        IReadOnlyList<string> header,
        IReadOnlyDictionary<string, CsvSchemaColumn>? schemaColumns)
    {
        var columns = new CsvDataColumnProjection[header.Count];
        for (var i = 0; i < columns.Length; i++)
        {
            var columnName = header[i];
            CsvSchemaColumn? schemaColumn = null;
            schemaColumns?.TryGetValue(columnName, out schemaColumn);
            columns[i] = new CsvDataColumnProjection(columnName, ResolveDataColumnType(schemaColumn?.DataType), schemaColumn);
        }

        return columns;
    }

    public static Type ResolveDataColumnType(Type? dataType)
    {
        if (dataType is null)
        {
            return typeof(string);
        }

        return Nullable.GetUnderlyingType(dataType) ?? dataType;
    }
}
