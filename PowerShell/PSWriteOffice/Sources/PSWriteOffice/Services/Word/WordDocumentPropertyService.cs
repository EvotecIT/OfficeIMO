using System;
using System.Collections.Generic;
using System.Globalization;
using System.Management.Automation;
using DocumentFormat.OpenXml.CustomProperties;
using OfficeIMO.Word;
using PSWriteOffice.Models.Word;

namespace PSWriteOffice.Services.Word;

internal static class WordDocumentPropertyService
{
    private static readonly IReadOnlyDictionary<string, Func<WordBuiltinDocumentProperties, object?>> BuiltInReaders =
        new Dictionary<string, Func<WordBuiltinDocumentProperties, object?>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Title"] = properties => properties.Title,
            ["Subject"] = properties => properties.Subject,
            ["Creator"] = properties => properties.Creator,
            ["Keywords"] = properties => properties.Keywords,
            ["Description"] = properties => properties.Description,
            ["Category"] = properties => properties.Category,
            ["Revision"] = properties => properties.Revision,
            ["LastModifiedBy"] = properties => properties.LastModifiedBy,
            ["Version"] = properties => properties.Version,
            ["Created"] = properties => properties.Created,
            ["Modified"] = properties => properties.Modified,
            ["LastPrinted"] = properties => properties.LastPrinted
        };

    private static readonly IReadOnlyDictionary<string, Action<WordBuiltinDocumentProperties, object?>> BuiltInWriters =
        new Dictionary<string, Action<WordBuiltinDocumentProperties, object?>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Title"] = (properties, value) => properties.Title = ConvertToString(value),
            ["Subject"] = (properties, value) => properties.Subject = ConvertToString(value),
            ["Creator"] = (properties, value) => properties.Creator = ConvertToString(value),
            ["Keywords"] = (properties, value) => properties.Keywords = ConvertToString(value),
            ["Description"] = (properties, value) => properties.Description = ConvertToString(value),
            ["Category"] = (properties, value) => properties.Category = ConvertToString(value),
            ["Revision"] = (properties, value) => properties.Revision = ConvertToString(value),
            ["LastModifiedBy"] = (properties, value) => properties.LastModifiedBy = ConvertToString(value),
            ["Version"] = (properties, value) => properties.Version = ConvertToString(value),
            ["Created"] = (properties, value) => properties.Created = ConvertToDateTime(value),
            ["Modified"] = (properties, value) => properties.Modified = ConvertToDateTime(value),
            ["LastPrinted"] = (properties, value) => properties.LastPrinted = ConvertToDateTime(value)
        };

    public static IEnumerable<WordDocumentPropertyInfo> GetProperties(WordDocument document, bool includeBuiltIn, bool includeCustom)
    {
        if (includeBuiltIn)
        {
            foreach (var property in BuiltInReaders)
            {
                var value = property.Value(document.BuiltinDocumentProperties);
                yield return new WordDocumentPropertyInfo(
                    property.Key,
                    "BuiltIn",
                    value,
                    value?.GetType().FullName,
                    null);
            }
        }

        if (includeCustom)
        {
            foreach (var property in document.CustomDocumentProperties)
            {
                var value = property.Value.Value;
                yield return new WordDocumentPropertyInfo(
                    property.Key,
                    "Custom",
                    value,
                    value?.GetType().FullName,
                    property.Value.PropertyType.ToString());
            }
        }
    }

    public static bool IsBuiltInProperty(string name)
    {
        return BuiltInWriters.ContainsKey(name);
    }

    public static void SetBuiltInProperty(WordDocument document, string name, object? value)
    {
        if (!BuiltInWriters.TryGetValue(name, out var writer))
        {
            throw new PSArgumentException(
                $"'{name}' is not a supported built-in document property. Use -Custom for custom properties.");
        }

        writer(document.BuiltinDocumentProperties, UnwrapValue(value));
    }

    public static void SetCustomProperty(WordDocument document, string name, object? value)
    {
        document.CustomDocumentProperties[name] = CreateCustomProperty(value);
    }

    private static WordCustomProperty CreateCustomProperty(object? value)
    {
        value = UnwrapValue(value);
        if (value == null)
        {
            return new WordCustomProperty(string.Empty);
        }

        if (value is bool boolean)
        {
            return new WordCustomProperty(boolean);
        }

        if (value is DateTime dateTime)
        {
            return new WordCustomProperty(dateTime);
        }

        if (value is DateTimeOffset dateTimeOffset)
        {
            return new WordCustomProperty(dateTimeOffset.UtcDateTime);
        }

        if (TryConvertToInt32(value, out var integer))
        {
            return new WordCustomProperty(integer);
        }

        if (TryConvertToDouble(value, out var number))
        {
            return new WordCustomProperty(number);
        }

        return new WordCustomProperty(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
    }

    private static object? UnwrapValue(object? value)
    {
        return value is PSObject psObject ? psObject.BaseObject : value;
    }

    private static string? ConvertToString(object? value)
    {
        value = UnwrapValue(value);
        return value == null ? null : LanguagePrimitives.ConvertTo<string>(value);
    }

    private static DateTime? ConvertToDateTime(object? value)
    {
        value = UnwrapValue(value);
        if (value == null)
        {
            return null;
        }

        if (value is DateTime dateTime)
        {
            return dateTime;
        }

        if (value is DateTimeOffset dateTimeOffset)
        {
            return dateTimeOffset.DateTime;
        }

        return LanguagePrimitives.ConvertTo<DateTime>(value);
    }

    private static bool TryConvertToInt32(object value, out int result)
    {
        switch (value)
        {
            case byte byteValue:
                result = byteValue;
                return true;
            case sbyte sbyteValue:
                result = sbyteValue;
                return true;
            case short shortValue:
                result = shortValue;
                return true;
            case ushort ushortValue:
                result = ushortValue;
                return true;
            case int intValue:
                result = intValue;
                return true;
            case long longValue when longValue >= int.MinValue && longValue <= int.MaxValue:
                result = (int)longValue;
                return true;
            case uint uintValue when uintValue <= int.MaxValue:
                result = (int)uintValue;
                return true;
            case ulong ulongValue when ulongValue <= int.MaxValue:
                result = (int)ulongValue;
                return true;
            default:
                result = default;
                return false;
        }
    }

    private static bool TryConvertToDouble(object value, out double result)
    {
        switch (value)
        {
            case float floatValue:
                result = floatValue;
                return true;
            case double doubleValue:
                result = doubleValue;
                return true;
            case decimal decimalValue:
                result = (double)decimalValue;
                return true;
            default:
                result = default;
                return false;
        }
    }
}
