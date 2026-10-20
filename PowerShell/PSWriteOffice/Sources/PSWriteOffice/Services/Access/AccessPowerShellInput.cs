using System;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Access;

namespace PSWriteOffice.Services.Access;

/// <summary>Applies PowerShell scalar conversion at the native table's declared type boundary.</summary>
internal static class AccessPowerShellInput {
    internal static AccessRowValues ToRow(AccessTable table, PSObject input) {
        var result = new AccessRowValues();
        if (PowerShellDictionaryAdapter.TryGetRowEntries(input, out var entries)) {
            foreach (var entry in entries) {
                string name = LanguagePrimitives.ConvertTo<string>(entry.Key);
                result.Add(name, ConvertValue(table.Columns[name], entry.Value));
            }
        } else {
            foreach (var property in input.Properties) {
                if (property.IsGettable) result.Add(property.Name, ConvertValue(table.Columns[property.Name], property.Value));
            }
        }
        return result;
    }

    private static object? ConvertValue(AccessColumn column, object? value) {
        if (value is PSObject wrapped) value = wrapped.BaseObject;
        if (value == null || value == DBNull.Value) return null;
        if (column.DataType == AccessDataType.Boolean && value is string text) {
            if (bool.TryParse(text, out bool boolean)) return boolean;
            if (text == "0") return false;
            if (text == "1") return true;
            throw new PSArgumentException($"Field '{column.Name}' requires True, False, 0, or 1.");
        }
        Type? type = column.DataType switch {
            AccessDataType.AutoNumber or AccessDataType.Int32 => typeof(int),
            AccessDataType.Int16 => typeof(short),
            AccessDataType.Int64 => typeof(long),
            AccessDataType.Byte => typeof(byte),
            AccessDataType.ShortText or AccessDataType.LongText => typeof(string),
            AccessDataType.Currency or AccessDataType.Decimal => typeof(decimal),
            AccessDataType.Double => typeof(double),
            AccessDataType.Single => typeof(float),
            AccessDataType.Boolean => typeof(bool),
            AccessDataType.DateTime => typeof(DateTime),
            AccessDataType.Guid => typeof(Guid),
            AccessDataType.Binary => typeof(byte[]),
            _ => null
        };
        return type == null ? value : LanguagePrimitives.ConvertTo(value, type, CultureInfo.InvariantCulture);
    }
}
