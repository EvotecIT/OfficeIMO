using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq.Expressions;
using System.Management.Automation;
using System.Reflection;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

internal sealed class CsvPowerShellRowWriter
{
    private bool _prevalidatedOutputProperties;
    private string[]? _outputHeader;

    public void Reset()
    {
        _outputHeader = null;
        _prevalidatedOutputProperties = false;
    }

    public void WriteDocumentRows(CsvDocument document, bool asHashtable, PSCmdlet cmdlet)
    {
        var header = document.Header;
        foreach (var row in document.AsEnumerable())
        {
            if (asHashtable)
            {
                var rowValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (var i = 0; i < header.Count && i < row.FieldCount; i++)
                {
                    rowValues.Add(header[i], row[i]);
                }

                cmdlet.WriteObject(rowValues);
                continue;
            }

            var outputHeader = GetOutputHeader(header);
            var headerCount = outputHeader.Length;
            var valueCount = headerCount < row.FieldCount ? headerCount : row.FieldCount;
            var psObj = PowerShellObjectFactory.Create(headerCount);
            var prevalidated = _prevalidatedOutputProperties;
            for (var i = 0; i < valueCount; i++)
            {
                psObj.Properties.Add(new PSNoteProperty(outputHeader[i], row[i]), prevalidated);
            }

            if (valueCount < headerCount)
            {
                AddMissingNoteProperties(psObj, outputHeader, valueCount, headerCount, prevalidated);
            }

            cmdlet.WriteObject(psObj);
        }
    }

    public void WriteDataReaderRows(DbDataReader reader, bool asHashtable, PSCmdlet cmdlet)
    {
        if (reader == null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        var header = new string[reader.FieldCount];
        for (var index = 0; index < header.Length; index++)
        {
            header[index] = reader.GetName(index);
        }

        var outputHeader = GetOutputHeader(header);
        var prevalidated = _prevalidatedOutputProperties;
        while (reader.Read())
        {
            if (asHashtable)
            {
                var values = new Dictionary<string, object?>(outputHeader.Length, StringComparer.OrdinalIgnoreCase);
                for (var index = 0; index < outputHeader.Length; index++)
                {
                    var value = reader.GetValue(index);
                    values.Add(outputHeader[index], value is DBNull ? null : value);
                }

                cmdlet.WriteObject(values);
                continue;
            }

            var row = PowerShellObjectFactory.Create(outputHeader.Length);
            for (var index = 0; index < outputHeader.Length; index++)
            {
                var value = reader.GetValue(index);
                row.Properties.Add(new PSNoteProperty(outputHeader[index], value is DBNull ? null : value), prevalidated);
            }

            cmdlet.WriteObject(row);
        }
    }

    public void WriteRow(IReadOnlyList<string> header, IReadOnlyList<string> row, bool asHashtable, PSCmdlet cmdlet)
    {
        var outputHeader = GetOutputHeader(header);
        var headerCount = outputHeader.Length;

        if (asHashtable)
        {
            var valueCount = GetValueCount(row, headerCount);
            var rowValues = new Dictionary<string, object?>(valueCount, StringComparer.OrdinalIgnoreCase);
            AddHashtableValues(rowValues, outputHeader, row, valueCount);

            cmdlet.WriteObject(rowValues);
            return;
        }

        WriteObjectRow(cmdlet, outputHeader, row, headerCount);
    }

    private void WriteObjectRow(PSCmdlet cmdlet, string[] header, IReadOnlyList<string> row, int headerCount)
    {
        var valueCount = GetValueCount(row, headerCount);
        var psObj = PowerShellObjectFactory.Create(headerCount);
        var prevalidated = _prevalidatedOutputProperties;

        if (row is List<string> list)
        {
            for (var i = 0; i < valueCount; i++)
            {
                psObj.Properties.Add(new PSNoteProperty(header[i], list[i]), prevalidated);
            }

            if (valueCount < headerCount)
            {
                AddMissingNoteProperties(psObj, header, valueCount, headerCount, prevalidated);
            }

            cmdlet.WriteObject(psObj);
            return;
        }

        if (row is string[] array)
        {
            for (var i = 0; i < valueCount; i++)
            {
                psObj.Properties.Add(new PSNoteProperty(header[i], array[i]), prevalidated);
            }

            if (valueCount < headerCount)
            {
                AddMissingNoteProperties(psObj, header, valueCount, headerCount, prevalidated);
            }

            cmdlet.WriteObject(psObj);
            return;
        }

        for (var i = 0; i < valueCount; i++)
        {
            psObj.Properties.Add(new PSNoteProperty(header[i], row[i]), prevalidated);
        }

        if (valueCount < headerCount)
        {
            AddMissingNoteProperties(psObj, header, valueCount, headerCount, prevalidated);
        }

        cmdlet.WriteObject(psObj);
    }

    private static int GetValueCount(IReadOnlyCollection<string> row, int headerCount) =>
        row.Count < headerCount ? row.Count : headerCount;

    private static int GetValueCount(IReadOnlyList<string> row, int headerCount) =>
        row.Count < headerCount ? row.Count : headerCount;

    private static void AddHashtableValues(Dictionary<string, object?> rowValues, string[] header, IReadOnlyList<string> row, int valueCount)
    {
        if (row is List<string> list)
        {
            for (var i = 0; i < valueCount; i++)
            {
                rowValues.Add(header[i], list[i]);
            }

            return;
        }

        if (row is string[] array)
        {
            for (var i = 0; i < valueCount; i++)
            {
                rowValues.Add(header[i], array[i]);
            }

            return;
        }

        for (var i = 0; i < valueCount; i++)
        {
            rowValues.Add(header[i], row[i]);
        }
    }

    internal static void AddMissingNoteProperties(PSObject psObj, string[] header, int startIndex, int headerCount, bool prevalidated)
    {
        for (var i = startIndex; i < headerCount; i++)
        {
            psObj.Properties.Add(new PSNoteProperty(header[i], null), prevalidated);
        }
    }

    internal bool PrevalidatedOutputProperties => _prevalidatedOutputProperties;

    internal string[] GetOutputHeader(IReadOnlyList<string> header)
    {
        if (_outputHeader is not null)
        {
            return _outputHeader;
        }

        var outputHeader = new string[header.Count];
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var canPrevalidate = true;
        for (var i = 0; i < header.Count; i++)
        {
            var name = header[i] ?? string.Empty;
            outputHeader[i] = name;

            if (string.IsNullOrEmpty(name))
            {
                canPrevalidate = false;
                continue;
            }

            if (!seen.Add(name))
            {
                throw new ExtendedTypeSystemException($"The member '{name}' is already present.");
            }
        }

        _prevalidatedOutputProperties = canPrevalidate;
        _outputHeader = outputHeader;
        return outputHeader;
    }

    internal static class PowerShellObjectFactory
    {
        // PowerShell 7+ exposes PSObject(int) for initial member capacity, but PowerShellStandard also exposes
        // PSObject(object). Use reflection so older runtimes fall back without wrapping the integer as BaseObject.
        private static readonly Func<int, PSObject>? CapacityFactory = CreateCapacityFactory();

        public static PSObject Create(int capacity)
        {
            return CapacityFactory?.Invoke(capacity) ?? new PSObject();
        }

        private static Func<int, PSObject>? CreateCapacityFactory()
        {
            ConstructorInfo? constructor = null;
            foreach (var candidate in typeof(PSObject).GetConstructors(BindingFlags.Public | BindingFlags.Instance))
            {
                var parameters = candidate.GetParameters();
                if (parameters.Length == 1 && parameters[0].ParameterType == typeof(int))
                {
                    constructor = candidate;
                    break;
                }
            }

            if (constructor is null)
            {
                return null;
            }

            try
            {
                var capacity = Expression.Parameter(typeof(int), "capacity");
                return Expression.Lambda<Func<int, PSObject>>(
                    Expression.New(constructor, capacity),
                    capacity).Compile();
            }
            catch
            {
                return capacity => (PSObject)constructor.Invoke(new object[] { capacity });
            }
        }
    }
}
