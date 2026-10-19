using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelTemplateValueService
{
    public static Dictionary<string, object?> ConvertValues(Hashtable values)
    {
        var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (DictionaryEntry entry in values)
        {
            var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            if (!string.IsNullOrWhiteSpace(key))
            {
                result[key!] = entry.Value;
            }
        }

        return result;
    }

    public static IReadOnlyList<IDictionary<string, object?>> ConvertRows(IEnumerable<object?> rows)
    {
        return rows.Select(ConvertRow).ToArray();
    }

    public static ExcelTemplateOptions CreateOptions(string? cultureName, ExcelTemplateMissingValueBehavior? missingValueBehavior, bool throwOnMissing)
    {
        var options = new ExcelTemplateOptions
        {
            ThrowOnMissing = throwOnMissing
        };

        if (missingValueBehavior.HasValue)
        {
            options.MissingValueBehavior = missingValueBehavior.Value;
        }

        if (!string.IsNullOrWhiteSpace(cultureName))
        {
            options.FormatProvider = CultureInfo.GetCultureInfo(cultureName!);
        }

        return options;
    }

    public static string? GetStringValue(IDictionary<string, object?> values, string? propertyName)
    {
        if (string.IsNullOrWhiteSpace(propertyName)
            || !values.TryGetValue(propertyName!, out var value)
            || value == null)
        {
            return null;
        }

        return Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    public static PSObject CreateMarkerRecord(ExcelTemplateMarkerInfo marker, string? path)
    {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("SheetName", marker.SheetName));
        record.Properties.Add(new PSNoteProperty("Sheet", marker.SheetName));
        record.Properties.Add(new PSNoteProperty("Address", marker.CellReference));
        record.Properties.Add(new PSNoteProperty("CellReference", marker.CellReference));
        record.Properties.Add(new PSNoteProperty("Name", marker.Name));
        record.Properties.Add(new PSNoteProperty("Marker", marker.Name));
        record.Properties.Add(new PSNoteProperty("Format", marker.Format));
        record.Properties.Add(new PSNoteProperty("CellText", marker.CellText));
        record.Properties.Add(new PSNoteProperty("IsWholeCell", marker.IsWholeCell));
        record.Properties.Add(new PSNoteProperty("IsBound", marker.IsBound));
        record.Properties.Add(new PSNoteProperty("BoundValueKind", marker.BoundValueKind));
        record.Properties.Add(new PSNoteProperty("BoundValueTypeName", marker.BoundValueTypeName));
        if (!string.IsNullOrWhiteSpace(path))
        {
            record.Properties.Add(new PSNoteProperty("Path", path));
            record.Properties.Add(new PSNoteProperty("InputPath", path));
        }

        return record;
    }

    private static IDictionary<string, object?> ConvertRow(object? row)
    {
        if (row == null)
        {
            return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        }

        if (row is PSObject psRow
            && psRow.BaseObject != null
            && !ReferenceEquals(psRow.BaseObject, row)
            && psRow.BaseObject is not PSCustomObject)
        {
            return ConvertRow(psRow.BaseObject);
        }

        if (row is Hashtable hashtable)
        {
            return ConvertValues(hashtable);
        }

        if (row is IDictionary dictionary)
        {
            var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(key))
                {
                    result[key!] = entry.Value;
                }
            }

            return result;
        }

        if (TryConvertGenericDictionary(row, out var genericDictionary))
        {
            return genericDictionary;
        }

        if (TryConvertKeyValueEntries(row, out var keyValueEntries))
        {
            return keyValueEntries;
        }

        var psObject = PSObject.AsPSObject(row);
        var properties = psObject.Properties
            .Where(property => property.IsGettable && property.MemberType is PSMemberTypes.NoteProperty or PSMemberTypes.Property)
            .ToArray();
        var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in properties)
        {
            values[property.Name] = property.Value;
        }

        return values;
    }

    private static bool TryConvertGenericDictionary(object row, out IDictionary<string, object?> values)
    {
        values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        Type rowType = row.GetType();
        bool isStringKeyDictionary = rowType
            .GetInterfaces()
            .Concat(new[] { rowType })
            .Any(type => type.IsGenericType
                && (type.GetGenericTypeDefinition() == typeof(IDictionary<,>)
                    || type.GetGenericTypeDefinition() == typeof(IReadOnlyDictionary<,>))
                && type.GetGenericArguments()[0] == typeof(string));

        if (!isStringKeyDictionary || row is not IEnumerable entries)
        {
            return false;
        }

        foreach (var entry in entries)
        {
            if (entry == null)
            {
                continue;
            }

            Type entryType = entry.GetType();
            var keyProperty = entryType.GetProperty("Key");
            var valueProperty = entryType.GetProperty("Value");
            if (keyProperty == null || valueProperty == null)
            {
                continue;
            }

            var key = Convert.ToString(keyProperty.GetValue(entry), CultureInfo.InvariantCulture);
            if (!string.IsNullOrWhiteSpace(key))
            {
                values[key!] = valueProperty.GetValue(entry);
            }
        }

        return true;
    }

    private static bool TryConvertKeyValueEntries(object row, out IDictionary<string, object?> values)
    {
        values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        if (row is string || row is not IEnumerable entries)
        {
            return false;
        }

        bool hasEntries = false;
        foreach (var entry in entries)
        {
            if (entry == null)
            {
                return false;
            }

            Type entryType = entry.GetType();
            var keyProperty = entryType.GetProperty("Key");
            var valueProperty = entryType.GetProperty("Value");
            if (keyProperty == null || valueProperty == null)
            {
                return false;
            }

            var key = Convert.ToString(keyProperty.GetValue(entry), CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(key))
            {
                return false;
            }

            values[key!] = valueProperty.GetValue(entry);
            hasEntries = true;
        }

        return hasEntries;
    }
}
