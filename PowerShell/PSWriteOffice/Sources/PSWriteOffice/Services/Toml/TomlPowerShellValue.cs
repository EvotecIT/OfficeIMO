using System.Collections;
using System.Collections.Specialized;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Toml;

namespace PSWriteOffice.Services.Toml;

/// <summary>
/// Moves values between PowerShell and TOML documents. Reading walks the document model: tables become
/// PSCustomObjects (or ordered hashtables), arrays and arrays of tables become arrays, and date-times become
/// DateTimeOffset, DateTime or TimeSpan by their TOML kind. Writing projects PowerShell objects the way ConvertTo-Json
/// reads them and builds a document whose nested objects are <c>[table]</c> sections and whose lists of objects are
/// <c>[[array]]</c> sections.
/// </summary>
internal static class TomlPowerShellValue {
    /// <summary>The largest -Depth, matching ConvertTo-Json.</summary>
    internal const int MaxDepth = 100;

    /// <summary>Converts a TOML node to PowerShell values.</summary>
    internal static object? FromToml(TomlNode node, bool asHashtable) {
        switch (node) {
            case TomlTable table:
                if (asHashtable) {
                    var dictionary = new OrderedDictionary(table.Count);
                    foreach (KeyValuePair<string, TomlNode> entry in table) dictionary[entry.Key] = FromToml(entry.Value, true);
                    return dictionary;
                }

                var result = new PSObject();
                foreach (KeyValuePair<string, TomlNode> entry in table) result.Properties.Add(new PSNoteProperty(entry.Key, FromToml(entry.Value, false)));
                return result;
            case TomlTableArray tables:
                var rows = new object?[tables.Count];
                for (int i = 0; i < rows.Length; i++) rows[i] = FromToml(tables[i], asHashtable);
                return rows;
            case TomlArray array:
                var items = new object?[array.Count];
                for (int i = 0; i < items.Length; i++) items[i] = FromToml(array[i], asHashtable);
                return items;
            case TomlValue value when value.Value is TomlDateTime dateTime:
                return dateTime.Kind switch {
                    TomlDateTimeKind.OffsetDateTime => dateTime.ToDateTimeOffset(),
                    TomlDateTimeKind.LocalTime => dateTime.ToTimeOfDay(),
                    _ => dateTime.ToDateTime()
                };
            case TomlValue value:
                return value.Value;
            default:
                return null;
        }
    }

    /// <summary>
    /// Builds a TOML document from a PowerShell value. A value that is not a dictionary or object is written under
    /// <paramref name="rootName"/>, because a TOML document is a table. TOML has no null, so nulls are omitted and
    /// counted; values deeper than <paramref name="depth"/> are written as their string form and set
    /// <paramref name="truncated"/>, matching ConvertTo-Json's -Depth.
    /// </summary>
    internal static TomlDocument ToDocument(object? value, int depth, string rootName, out bool wrapped, out int nulls, out bool truncated) {
        truncated = false;
        object? projected = Project(value, 0, depth, ref truncated);
        wrapped = projected is not Dictionary<string, object?>;
        var root = wrapped ? new Dictionary<string, object?> { [rootName] = projected } : (Dictionary<string, object?>)projected!;
        TomlDocument document = TomlDocument.Create();
        nulls = 0;
        WriteTable(document.Root, root, ref nulls);
        return document;
    }

    /// <summary>
    /// Converts a value for Set-OfficeTomlNode: nodes and scalars as they are, hashtables and objects as inline
    /// tables, and lists as arrays. Objects nest at most <see cref="MaxDepth"/> levels, and nested nulls are omitted
    /// and counted, because TOML has no null.
    /// </summary>
    /// <exception cref="ArgumentException">The value is null.</exception>
    internal static object ToValue(object? value, out int nulls, out bool truncated) {
        truncated = false;
        nulls = 0;
        object? projected = Project(value, 0, MaxDepth, ref truncated);
        if (projected is null) throw new ArgumentException("TOML has no null; use Remove-OfficeTomlNode to remove a value.");
        return WithoutNulls(projected, ref nulls)!;
    }

    // A table's values come before its sub-tables, as TOML requires.
    private static void WriteTable(TomlTable table, Dictionary<string, object?> members, ref int nulls) {
        foreach (KeyValuePair<string, object?> member in members) {
            if (member.Value is null) {
                nulls++;
                continue;
            }

            if (IsSection(member.Value)) continue;
            table.Set(member.Key, WithoutNulls(member.Value, ref nulls)!);
        }

        foreach (KeyValuePair<string, object?> member in members) {
            if (member.Value is Dictionary<string, object?> child) {
                WriteTable(table.AddTable(member.Key), child, ref nulls);
            } else if (member.Value is List<object?> items && IsSection(items)) {
                TomlTableArray array = table.AddTableArray(member.Key);
                foreach (object? item in items) WriteTable(array.AddTable(), (Dictionary<string, object?>)item!, ref nulls);
            }
        }
    }

    private static bool IsSection(object value) =>
        value is Dictionary<string, object?> || (value is List<object?> items && items.Count > 0 && items.TrueForAll(item => item is Dictionary<string, object?>));

    // Values on a key's line: nulls inside arrays and inline tables are dropped too.
    private static object? WithoutNulls(object? value, ref int nulls) {
        switch (value) {
            case Dictionary<string, object?> dictionary:
                var table = new Dictionary<string, object?>(dictionary.Count);
                foreach (KeyValuePair<string, object?> entry in dictionary) {
                    if (entry.Value is null) nulls++;
                    else table[entry.Key] = WithoutNulls(entry.Value, ref nulls);
                }

                return table;
            case List<object?> list:
                var items = new List<object?>(list.Count);
                foreach (object? item in list) {
                    if (item is null) nulls++;
                    else items.Add(WithoutNulls(item, ref nulls));
                }

                return items;
            default:
                return value;
        }
    }

    private static object? Project(object? value, int level, int depth, ref bool truncated) {
        object? item = Unwrap(value);
        switch (item) {
            case null: return null;
            case string text: return text;
            case char character: return character.ToString();
            case bool flag: return flag;
            case sbyte or byte or short or ushort or int or uint or long: return Convert.ToInt64(item, CultureInfo.InvariantCulture);
            case ulong unsigned: return unsigned <= long.MaxValue ? (long)unsigned : unsigned.ToString(CultureInfo.InvariantCulture);
            case float or double or decimal: return Convert.ToDouble(item, CultureInfo.InvariantCulture);
            case System.Numerics.BigInteger big: return big >= long.MinValue && big <= long.MaxValue ? (long)big : big.ToString(CultureInfo.InvariantCulture);
            case DateTime or DateTimeOffset or TomlDateTime: return item;
            case Enum enumeration: return enumeration.ToString();
            case Guid or Uri or TimeSpan or Version: return Convert.ToString(item, CultureInfo.InvariantCulture);
            case ScriptBlock script: return script.ToString();
        }

        if (level > depth) {
            truncated = true;
            return Convert.ToString(item, CultureInfo.InvariantCulture);
        }

        if (item is IDictionary dictionary) {
            var map = new Dictionary<string, object?>(dictionary.Count);
            foreach (DictionaryEntry entry in dictionary) {
                map[Convert.ToString(Unwrap(entry.Key), CultureInfo.InvariantCulture) ?? string.Empty] = Project(entry.Value, level + 1, depth, ref truncated);
            }

            return map;
        }

        if (item is IEnumerable sequence) {
            var list = new List<object?>();
            foreach (object? element in sequence) list.Add(Project(element, level + 1, depth, ref truncated));
            return list;
        }

        var properties = new Dictionary<string, object?>();
        foreach (PSPropertyInfo property in PSObject.AsPSObject(value!).Properties) {
            if (!property.IsGettable) continue;
            object? propertyValue;
            try { propertyValue = property.Value; }
            catch (GetValueException) { continue; }
            properties[property.Name] = Project(propertyValue, level + 1, depth, ref truncated);
        }

        return properties;
    }

    private static object? Unwrap(object? value) {
        // A PSCustomObject's base object is a placeholder; its note properties are the value.
        while (value is PSObject wrapped && wrapped.BaseObject is not PSCustomObject) value = wrapped.BaseObject;
        return value;
    }
}
