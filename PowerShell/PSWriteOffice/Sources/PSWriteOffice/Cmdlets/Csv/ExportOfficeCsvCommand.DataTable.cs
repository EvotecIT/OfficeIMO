using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

public sealed partial class ExportOfficeCsvCommand
{
    private void ExportDataTable(DataTable table)
    {
        if (table == null)
        {
            return;
        }

        var sourceColumns = GetDataTableColumnNames(table);
        var hadActiveWriter = _streamingWriter != null;
        var writer = EnsureStreamingWriterForColumns(
            sourceColumns,
            out var effectiveColumns,
            columns => ValidateDataTableAppendColumns(table, sourceColumns, columns));
        if (writer == null)
        {
            return;
        }

        try
        {
            if (hadActiveWriter)
            {
                ValidateDataTableAppendColumns(table, sourceColumns, effectiveColumns);
            }

            if (!hadActiveWriter && ColumnsMatch(sourceColumns, effectiveColumns))
            {
                using var reader = table.CreateDataReader();
                writer.WriteDataReader(_objectProjector.NormalizeDataReader(reader));
            }
            else
            {
                WriteDataTableRows(table, writer, effectiveColumns);
            }
        }
        catch
        {
            DisposeStreamingWriter();
            throw;
        }
    }

    private static bool TryGetDataTable(object? value, out DataTable table)
    {
        if (value is DataTable dataTable)
        {
            table = dataTable;
            return true;
        }

        if (value is PSObject { BaseObject: DataTable psObjectTable })
        {
            table = psObjectTable;
            return true;
        }

        table = null!;
        return false;
    }

    private static bool TryGetDataView(object? value, out DataView view)
    {
        if (value is DataView dataView)
        {
            view = dataView;
            return true;
        }

        if (value is PSObject { BaseObject: DataView psObjectView })
        {
            view = psObjectView;
            return true;
        }

        view = null!;
        return false;
    }

    private void AppendDataTable(DataTable table)
    {
        var options = CreateSaveOptions(includeHeader: !NoHeader.IsPresent && !_appendToExistingFile);
        _objectProjector.UseCsvOptions(options, MaxCollectionItems, MaxNestingDepth);
        var appendHeader = GetEffectiveAppendHeader(table);

        if (appendHeader is { Length: > 0 })
        {
            ValidateDataTableAppendHeader(table, appendHeader);
        }

        using var csvWriter = CreateRowWriter(append: true, options);
        if (appendHeader is { Length: > 0 })
        {
            WriteDataTableRows(table, csvWriter, appendHeader);
            return;
        }

        using var reader = table.CreateDataReader();
        csvWriter.WriteDataReader(_objectProjector.NormalizeDataReader(reader));
    }

    private static string[] GetDataTableColumnNames(DataTable table)
    {
        var columns = new string[table.Columns.Count];
        for (var i = 0; i < columns.Length; i++)
        {
            columns[i] = table.Columns[i].ColumnName;
        }

        return columns;
    }

    private string[]? GetEffectiveAppendHeader(DataTable table)
    {
        if (_appendHeader is not { Length: > 0 })
        {
            return null;
        }

        return _appendHeader;
    }

    private void ValidateDataTableAppendHeader(DataTable table, IReadOnlyList<string> appendHeader)
    {
        if (Force.IsPresent)
        {
            return;
        }

        foreach (var column in appendHeader)
        {
            if (!ContainsDataColumn(table, column))
            {
                throw new CsvException($"Cannot append CSV because the DataTable is missing the existing column '{column}'. Use -Force to append with blank values for missing columns.");
            }
        }
    }

    private void ValidateDataTableAppendColumns(DataTable table, IReadOnlyList<string> sourceColumns, IReadOnlyList<string> effectiveColumns)
    {
        if (Append.IsPresent &&
            effectiveColumns.Count > 0 &&
            !ColumnsMatch(sourceColumns, effectiveColumns))
        {
            ValidateDataTableAppendHeader(table, effectiveColumns);
        }
    }

    private void WriteDataTableRows(DataTable table, CsvRowWriter writer, IReadOnlyList<string> columns)
    {
        foreach (DataRow row in table.Rows)
        {
            writer.WriteRow(
                columns,
                columns.Count,
                (Row: row, Columns: columns, Projector: _objectProjector),
                static (state, index) => state.Projector.NormalizeCellValue(
                    TryGetDataTableValue(state.Row, state.Columns[index])));
        }
    }

    private static bool ColumnsMatch(IReadOnlyList<string> left, IReadOnlyList<string> right)
    {
        if (left.Count != right.Count)
        {
            return false;
        }

        for (var i = 0; i < left.Count; i++)
        {
            if (!string.Equals(left[i], right[i], StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        return true;
    }

    private static object? TryGetDataTableValue(DataRow row, string column)
    {
        var dataColumn = GetDataColumn(row.Table, column);
        if (dataColumn == null)
        {
            return null;
        }

        var value = row[dataColumn];
        return value == DBNull.Value ? null : value;
    }

    private static bool ContainsDataColumn(DataTable table, string column) => GetDataColumn(table, column) != null;

    private static DataColumn? GetDataColumn(DataTable table, string column)
    {
        if (table.Columns.Contains(column))
        {
            return table.Columns[column];
        }

        foreach (DataColumn dataColumn in table.Columns)
        {
            if (string.Equals(dataColumn.ColumnName, column, StringComparison.OrdinalIgnoreCase))
            {
                return dataColumn;
            }
        }

        return null;
    }
}
