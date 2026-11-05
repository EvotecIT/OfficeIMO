using System;
using System.Collections;
using System.Data;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

public sealed partial class ImportOfficeCsvCommand {
    private CsvDataTableOptions CreateDataTableOptions(string? tableName) =>
        new()
        {
            TableName = tableName,
            Schema = CreateExplicitSchema(),
            InferSchema = InferSchema.IsPresent,
            SchemaSampleRows = SchemaSampleRows
        };

    private CsvReaderOptions CreateDataReaderOptions() => ConfigureReaderOptions(new CsvReaderOptions());

    private CsvReaderOptions ConfigureReaderOptions(CsvReaderOptions options)
    {
        options.Schema = CreateExplicitSchema();
        options.InferSchema = InferSchema.IsPresent;
        options.SchemaSampleRows = SchemaSampleRows;
        return options;
    }

    private CsvSchema? CreateExplicitSchema()
    {
        if (ColumnType is not { Count: > 0 })
        {
            return null;
        }

        var builder = new CsvSchemaBuilder();
        foreach (DictionaryEntry entry in ColumnType)
        {
            var columnName = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(columnName))
            {
                throw new PSArgumentException("-ColumnType keys must be non-empty column names.");
            }

            builder.Column(columnName).AsType(ResolveColumnType(entry.Value, columnName));
        }

        return builder.Build();
    }

    private static Type ResolveColumnType(object? value, string columnName)
    {
        if (value is PSObject psObject)
        {
            value = psObject.BaseObject;
        }

        if (value is Type type)
        {
            return Nullable.GetUnderlyingType(type) ?? type;
        }

        if (value is string typeName)
        {
            return ResolveColumnTypeName(typeName, columnName);
        }

        throw new PSArgumentException($"ColumnType value for '{columnName}' must be a .NET type or a supported type name.");
    }

    private static Type ResolveColumnTypeName(string typeName, string columnName)
    {
        if (string.IsNullOrWhiteSpace(typeName))
        {
            throw new PSArgumentException($"ColumnType value for '{columnName}' cannot be empty.");
        }

        var trimmedTypeName = typeName.Trim();
        return trimmedTypeName.ToLowerInvariant() switch
        {
            "bool" or "boolean" => typeof(bool),
            "byte" => typeof(byte),
            "short" or "int16" => typeof(short),
            "int" or "int32" => typeof(int),
            "long" or "int64" => typeof(long),
            "decimal" => typeof(decimal),
            "double" => typeof(double),
            "float" or "single" => typeof(float),
            "datetime" => typeof(DateTime),
            "guid" => typeof(Guid),
            "string" => typeof(string),
            _ => Type.GetType(trimmedTypeName, throwOnError: false) ??
                throw new PSArgumentException($"ColumnType value '{typeName}' for '{columnName}' is not a supported type name.")
        };
    }

}
