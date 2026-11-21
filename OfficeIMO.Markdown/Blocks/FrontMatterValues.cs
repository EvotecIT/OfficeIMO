using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using OfficeIMO.Yaml;

namespace OfficeIMO.Markdown;

/// <summary>Converts .NET values supplied for new front matter into YAML nodes, which the YAML writer quotes safely.</summary>
internal static class FrontMatterValues {
    private const int MaxDepth = 64;

    /// <summary>Returns the mapping for a dictionary or key/value sequence, or <see langword="null"/> for other objects.</summary>
    public static YamlMapping? TryToMapping(object data, Func<object, int, YamlNode?>? complexValue) => TryToMapping(data, complexValue, depth: 0);

    [RequiresUnreferencedCode("Reads the public properties of arbitrary runtime types.")]
    public static YamlMapping ToMapping(object data, Func<object, int, YamlNode?> complexValue) =>
        TryToMapping(data, complexValue, depth: 0) ?? FromRuntimeProperties(data, depth: 0);

    public static YamlMapping FromProperties(object data,
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] Type type,
        Func<object, int, YamlNode?>? complexValue) {
        var mapping = new YamlMapping();
        foreach (PropertyInfo property in type.GetProperties()) {
            if (!IsReadable(property)) continue;
            mapping.Add(property.Name, ToNode(property.GetValue(data), complexValue, depth: 1));
        }

        return mapping;
    }

    /// <summary>Converts nested objects of any runtime type through their public properties.</summary>
    [RequiresUnreferencedCode("Reads the public properties of arbitrary runtime types.")]
    public static YamlNode? ReflectObject(object value, int depth) => FromRuntimeProperties(value, depth);

    [RequiresUnreferencedCode("Reads the public properties of arbitrary runtime types.")]
    private static YamlMapping FromRuntimeProperties(object data, int depth) {
        var mapping = new YamlMapping();
        foreach (PropertyInfo property in data.GetType().GetProperties()) {
            if (!IsReadable(property)) continue;
            mapping.Add(property.Name, ToNode(property.GetValue(data), ReflectObject, depth + 1));
        }

        return mapping;
    }

    private static bool IsReadable(PropertyInfo property) =>
        property.CanRead && property.GetIndexParameters().Length == 0 && property.GetMethod is { IsPublic: true };

    private static YamlMapping? TryToMapping(object data, Func<object, int, YamlNode?>? complexValue, int depth) {
        switch (data) {
            case YamlMapping mapping:
                return mapping;
            case IEnumerable<KeyValuePair<string, object?>> pairs: {
                var result = new YamlMapping();
                foreach (KeyValuePair<string, object?> pair in pairs) result.Add(pair.Key, ToNode(pair.Value, complexValue, depth + 1));
                return result;
            }
            case IDictionary dictionary: {
                var result = new YamlMapping();
                foreach (DictionaryEntry pair in dictionary) {
                    result.Add(Convert.ToString(pair.Key, CultureInfo.InvariantCulture) ?? string.Empty, ToNode(pair.Value, complexValue, depth + 1));
                }

                return result;
            }
            default:
                return null;
        }
    }

    private static YamlNode ToNode(object? value, Func<object, int, YamlNode?>? complexValue, int depth) {
        if (depth > MaxDepth) throw new InvalidOperationException("Front matter values nest deeper than " + MaxDepth.ToString(CultureInfo.InvariantCulture) + " levels; check for a reference cycle.");

        switch (value) {
            case null:
                return new YamlScalar((string?)null);
            case YamlNode node:
                return node;
            case string text:
                return new YamlScalar(text);
            case char character:
                return new YamlScalar(character.ToString());
            case bool boolean:
                return new YamlScalar(boolean);
            case Enum enumeration:
                return new YamlScalar(enumeration.ToString());
            case sbyte or byte or short or ushort or int or uint or long:
                return new YamlScalar(Convert.ToInt64(value, CultureInfo.InvariantCulture));
            case ulong unsigned:
                return new YamlScalar(unsigned.ToString(CultureInfo.InvariantCulture), YamlStyle.Plain);
            case float single:
                return new YamlScalar((double)single);
            case double number:
                return new YamlScalar(number);
            case decimal exact:
                return new YamlScalar(exact.ToString(CultureInfo.InvariantCulture), YamlStyle.Plain);
            case DateTime date:
                return new YamlScalar(FormatDate(date), YamlStyle.Plain);
            case DateTimeOffset dateOffset:
                return new YamlScalar(dateOffset.ToString("yyyy-MM-ddTHH:mm:ss.FFFFFFFzzz", CultureInfo.InvariantCulture), YamlStyle.Plain);
            case Uri uri:
                return new YamlScalar(uri.OriginalString);
            case IFormattable formattable when value is not IEnumerable:
                return new YamlScalar(formattable.ToString(null, CultureInfo.InvariantCulture));
        }

        YamlMapping? mapping = TryToMapping(value, complexValue, depth);
        if (mapping != null) return mapping;

        if (value is IEnumerable items) {
            var sequence = new YamlSequence();
            foreach (object? item in items) sequence.Add(ToNode(item, complexValue, depth + 1));
            return sequence;
        }

        return complexValue?.Invoke(value, depth) ?? new YamlScalar(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
    }

    private static string FormatDate(DateTime date) {
        if (date.TimeOfDay == TimeSpan.Zero && date.Kind != DateTimeKind.Utc) return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        return date.Kind == DateTimeKind.Utc
            ? date.ToString("yyyy-MM-ddTHH:mm:ss.FFFFFFFZ", CultureInfo.InvariantCulture)
            : date.ToString(date.Kind == DateTimeKind.Local ? "yyyy-MM-ddTHH:mm:ss.FFFFFFFzzz" : "yyyy-MM-ddTHH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture);
    }
}
