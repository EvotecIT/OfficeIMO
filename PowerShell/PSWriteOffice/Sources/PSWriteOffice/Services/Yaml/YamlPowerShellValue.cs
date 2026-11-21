using System.Collections;
using System.Collections.Specialized;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Yaml;

namespace PSWriteOffice.Services.Yaml;

/// <summary>
/// Moves values between PowerShell and the YAML serializer. Reading walks the serializer's untyped result
/// (dictionaries in file order, lists and schema-resolved scalars); writing projects PowerShell objects to the
/// dictionaries, lists and scalars the serializer writes without metadata, the way ConvertTo-Json reads them.
/// </summary>
internal static class YamlPowerShellValue {
    /// <summary>The largest -Depth, matching ConvertTo-Json.</summary>
    internal const int MaxDepth = 100;

    /// <summary>Converts an untyped serializer value to PSCustomObject (or ordered hashtable), object[] and scalars.</summary>
    internal static object? FromYaml(object? value, bool asHashtable) {
        switch (value) {
            case IDictionary dictionary:
                if (asHashtable) {
                    var table = new OrderedDictionary(dictionary.Count);
                    foreach (DictionaryEntry entry in dictionary) table[entry.Key] = FromYaml(entry.Value, true);
                    return table;
                }
                var result = new PSObject();
                foreach (DictionaryEntry entry in dictionary) {
                    string name = Convert.ToString(entry.Key, CultureInfo.InvariantCulture) ?? string.Empty;
                    // A repeated name can only come from keys that differ in type but print alike; the last one wins.
                    result.Properties.Remove(name);
                    result.Properties.Add(new PSNoteProperty(name, FromYaml(entry.Value, false)));
                }
                return result;
            case IList list:
                var items = new object?[list.Count];
                for (int i = 0; i < items.Length; i++) items[i] = FromYaml(list[i], asHashtable);
                return items;
            default:
                return value;
        }
    }

    /// <summary>
    /// Projects a PowerShell value for the serializer. Values deeper than <paramref name="depth"/> are written as their
    /// string form and <paramref name="truncated"/> is set, matching ConvertTo-Json's -Depth.
    /// </summary>
    internal static object? ToYaml(object? value, int depth, ref bool truncated) => Project(value, 0, depth, ref truncated);

    /// <summary>
    /// Converts a value for Set-OfficeYamlNode: nodes as they are, scalars directly, the rest through the serializer.
    /// Objects nest at most <see cref="MaxDepth"/> levels, as ConvertTo-Json's largest -Depth.
    /// </summary>
    internal static YamlNode ToNode(object? value, YamlStyle? style, out bool truncated) {
        truncated = false;
        object? unwrapped = Unwrap(value);
        if (unwrapped is YamlNode node) return node;
        if (style.HasValue) {
            if (unwrapped is not string && unwrapped is not char) throw new ArgumentException("Style applies to string values only.");
            return new YamlScalar(Convert.ToString(unwrapped, CultureInfo.InvariantCulture)!, style.Value);
        }
        object? projected = Project(value, 0, MaxDepth, ref truncated);
        return projected switch {
            null => new YamlScalar((string?)null),
            string text => new YamlScalar(text),
            bool flag => new YamlScalar(flag),
            long number => new YamlScalar(number),
            double real => new YamlScalar(real),
            _ => YamlSerializer.SerializeToNode(projected, projected.GetType())
        };
    }

    private static object? Project(object? value, int level, int depth, ref bool truncated) {
        object? item = Unwrap(value);
        switch (item) {
            case null: return null;
            case string text: return text;
            case char character: return character.ToString();
            case bool flag: return flag;
            case sbyte or byte or short or ushort or int or uint or long: return Convert.ToInt64(item, CultureInfo.InvariantCulture);
            case ulong unsigned: return unsigned <= long.MaxValue ? (long)unsigned : (object)unsigned;
            case float single: return (double)single;
            case double or decimal or System.Numerics.BigInteger or DateTime or DateTimeOffset: return item;
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
