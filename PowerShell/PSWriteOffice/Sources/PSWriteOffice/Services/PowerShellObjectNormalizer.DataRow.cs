using System;
using System.Collections.Generic;
using System.Data;
using System.Management.Automation;

namespace PSWriteOffice.Services;

/// <summary>Contains ADO.NET row projection for the shared PowerShell object normalizer.</summary>
internal static partial class PowerShellObjectNormalizer
{
    private static Dictionary<string, object?> ProjectDataRowDictionary(
        DataRow row,
        PowerShellObjectNormalizerOptions options)
    {
        var columns = row.Table.Columns;
        var result = new Dictionary<string, object?>(columns.Count, StringComparer.OrdinalIgnoreCase);
        foreach (DataColumn column in columns)
        {
            result[column.ColumnName] = NormalizeCellValue(row[column], options);
        }

        return result;
    }

    private static bool TryGetDataRow(object item, out DataRow row)
    {
        if (item is PSObject psObject)
        {
            item = psObject.BaseObject;
        }

        if (item is DataRow dataRow)
        {
            row = dataRow;
            return true;
        }

        if (item is DataRowView dataRowView)
        {
            row = dataRowView.Row;
            return true;
        }

        row = null!;
        return false;
    }

    private static void ProjectDataRow(
        DataRow row,
        string[]? columns,
        out string[] projectedColumns,
        out object?[] values,
        PowerShellObjectNormalizerOptions options)
    {
        if (columns == null)
        {
            var tableColumns = row.Table.Columns;
            projectedColumns = new string[tableColumns.Count];
            values = new object?[tableColumns.Count];
            for (var i = 0; i < tableColumns.Count; i++)
            {
                var column = tableColumns[i];
                projectedColumns[i] = column.ColumnName;
                values[i] = NormalizeCellValue(row[column], options);
            }

            return;
        }

        projectedColumns = columns;
        values = new object?[columns.Length];
        ProjectDataRowInto(row, columns, values, options);
    }

    private static void ProjectDataRowInto(
        DataRow row,
        string[] columns,
        object?[] values,
        PowerShellObjectNormalizerOptions options)
    {
        for (var i = 0; i < columns.Length; i++)
        {
            var column = GetDataColumn(row.Table.Columns, columns[i]);
            values[i] = column == null ? null : NormalizeCellValue(row[column], options);
        }
    }

    private static DataColumn? GetDataColumn(DataColumnCollection columns, string columnName)
    {
        if (columns.Contains(columnName))
        {
            return columns[columnName];
        }

        foreach (DataColumn column in columns)
        {
            if (string.Equals(column.ColumnName, columnName, StringComparison.OrdinalIgnoreCase))
            {
                return column;
            }
        }

        return null;
    }
}
