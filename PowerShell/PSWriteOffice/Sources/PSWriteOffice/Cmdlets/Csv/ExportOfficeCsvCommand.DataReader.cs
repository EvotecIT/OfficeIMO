using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

public sealed partial class ExportOfficeCsvCommand
{
    private void ExportDataReader(IDataReader reader)
    {
        if (reader == null)
        {
            return;
        }

        var allowAdditionalAppend = false;
        IReadOnlyList<string>? activeAppendColumns = null;
        if (Append.IsPresent && _streamingWriter != null && !IsCompressedAppendTarget(CompressionType, _resolvedPath!))
        {
            activeAppendColumns = NoHeader.IsPresent ? _objectProjector.CurrentColumns : null;
            DisposeStreamingWriter();
            allowAdditionalAppend = true;
        }

        var sourceColumns = GetDataReaderColumnNames(reader);
        var hadActiveWriter = _streamingWriter != null;
        var writer = EnsureStreamingWriterForColumns(
            sourceColumns,
            out var effectiveColumns,
            columns => ValidateDataReaderAppendColumns(reader, sourceColumns, columns),
            allowAdditionalAppend,
            activeAppendColumns);
        if (writer == null)
        {
            return;
        }

        try
        {
            if (hadActiveWriter)
            {
                ValidateDataReaderAppendColumns(reader, sourceColumns, effectiveColumns);
            }

            if (!hadActiveWriter && ColumnsMatch(sourceColumns, effectiveColumns))
            {
                writer.WriteDataReader(_objectProjector.NormalizeDataReader(reader));
            }
            else
            {
                WriteDataReaderRows(reader, writer, effectiveColumns);
            }
        }
        catch
        {
            DisposeStreamingWriter();
            throw;
        }
    }

    private static bool TryGetDataReader(object? value, out IDataReader reader)
    {
        if (value is IDataReader dataReader)
        {
            reader = dataReader;
            return true;
        }

        if (value is PSObject { BaseObject: IDataReader psObjectReader })
        {
            reader = psObjectReader;
            return true;
        }

        reader = null!;
        return false;
    }

    private void AppendDataReader(IDataReader reader)
    {
        var options = CreateSaveOptions(includeHeader: !NoHeader.IsPresent && !_appendToExistingFile);
        _objectProjector.UseCsvOptions(options, MaxCollectionItems, MaxNestingDepth);
        var appendHeader = GetEffectiveAppendHeader(reader);

        if (appendHeader is { Length: > 0 })
        {
            ValidateDataReaderAppendHeader(reader, appendHeader);
        }

        using var csvWriter = CreateRowWriter(append: true, options);
        if (appendHeader is { Length: > 0 })
        {
            WriteDataReaderRows(reader, csvWriter, appendHeader);
            return;
        }

        csvWriter.WriteDataReader(_objectProjector.NormalizeDataReader(reader));
    }

    private static string[] GetDataReaderColumnNames(IDataReader reader)
    {
        var columns = new string[reader.FieldCount];
        for (var i = 0; i < columns.Length; i++)
        {
            columns[i] = reader.GetName(i);
        }

        return columns;
    }

    private string[]? GetEffectiveAppendHeader(IDataReader reader)
    {
        if (_appendHeader is not { Length: > 0 })
        {
            return null;
        }

        return _appendHeader;
    }

    private void ValidateDataReaderAppendHeader(IDataReader reader, IReadOnlyList<string> appendHeader)
    {
        if (Force.IsPresent)
        {
            return;
        }

        foreach (var column in appendHeader)
        {
            if (!ContainsDataReaderColumn(reader, column))
            {
                throw new CsvException($"Cannot append CSV because the data reader is missing the existing column '{column}'. Use -Force to append with blank values for missing columns.");
            }
        }
    }

    private void ValidateDataReaderAppendColumns(IDataReader reader, IReadOnlyList<string> sourceColumns, IReadOnlyList<string> effectiveColumns)
    {
        if (Append.IsPresent &&
            effectiveColumns.Count > 0 &&
            !ColumnsMatch(sourceColumns, effectiveColumns))
        {
            ValidateDataReaderAppendHeader(reader, effectiveColumns);
        }
    }

    private void WriteDataReaderRows(IDataReader reader, CsvRowWriter writer, IReadOnlyList<string> columns)
    {
        var columnOrdinals = GetDataReaderColumnOrdinals(reader);
        while (reader.Read())
        {
            writer.WriteRow(
                columns,
                columns.Count,
                (Reader: reader, Columns: columns, Ordinals: columnOrdinals, Projector: _objectProjector),
                static (state, index) => state.Projector.NormalizeCellValue(
                    TryGetDataReaderValue(state.Reader, state.Columns[index], state.Ordinals)));
        }
    }

    private static object? TryGetDataReaderValue(IDataRecord reader, string column, IReadOnlyDictionary<string, int> ordinals)
    {
        if (!ordinals.TryGetValue(column, out var ordinal))
        {
            return null;
        }

        var value = reader.GetValue(ordinal);
        return value == DBNull.Value ? null : value;
    }

    private static bool ContainsDataReaderColumn(IDataReader reader, string column)
    {
        for (var i = 0; i < reader.FieldCount; i++)
        {
            if (string.Equals(reader.GetName(i), column, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static Dictionary<string, int> GetDataReaderColumnOrdinals(IDataReader reader)
    {
        var ordinals = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < reader.FieldCount; i++)
        {
            var name = reader.GetName(i);
            if (!ordinals.ContainsKey(name))
            {
                ordinals.Add(name, i);
            }
        }

        return ordinals;
    }
}
