using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;

namespace PSWriteOffice.Services.Text;

internal static class OfficeTextRunParser
{
    private static readonly string[][] SequenceProperties =
    {
        new[] { "Kind", "Type", "Run" },
        new[] { "Bold" },
        new[] { "Italic" },
        new[] { "Underline", "Underlined" },
        new[] { "UnderlineStyle", "UnderlineKind" },
        new[] { "Strike", "Strikethrough" },
        new[] { "Color", "TextColor", "FontColor" },
        new[] { "BackgroundColor", "HighlightColor", "FillColor" },
        new[] { "FontSize", "Size" },
        new[] { "FontName", "Font", "Typeface", "FontFamily" },
        new[] { "Baseline" },
        new[] { "LinkUri", "Uri", "Url", "Href" },
        new[] { "LinkDestinationName", "DestinationName", "Bookmark", "BookmarkName" },
        new[] { "LinkContents", "Contents", "Tooltip" },
        new[] { "TabLeader", "Leader" },
        new[] { "TabAlignment", "Alignment" }
    };

    internal static OfficeTextRunSpec[] ParseMany(object[]? runs)
    {
        var normalized = ToRunArray(runs);
        if (normalized.Length == 0)
        {
            throw new PSArgumentException("Provide at least one text run.");
        }

        return normalized.SelectMany(ExpandAndParse).ToArray();
    }

    internal static OfficeTextRunSpec[] ParseMany(object? runs)
    {
        var normalized = ToRunArray(runs);
        if (normalized.Length == 0)
        {
            throw new PSArgumentException("Provide at least one text run.");
        }

        return normalized.SelectMany(ExpandAndParse).ToArray();
    }

    internal static object[] ToRunArray(object? runs)
    {
        if (runs == null)
        {
            return Array.Empty<object>();
        }

        if (runs is string)
        {
            return new[] { runs };
        }

        if (runs is IDictionary || runs is OfficeTextRunSpec)
        {
            return new[] { runs };
        }

        return runs is IEnumerable enumerable
            ? enumerable.Cast<object>().ToArray()
            : new[] { runs };
    }

    private static IEnumerable<OfficeTextRunSpec> ExpandAndParse(object value)
    {
        var textValue = UnwrapPSObject(GetValue(value, "Text", "Value", "Content"));
        if (textValue is string || textValue is IDictionary || !(textValue is IEnumerable textValues))
        {
            yield return Parse(value);
            yield break;
        }

        var textItems = textValues.Cast<object?>().Select(UnwrapPSObject).ToArray();
        if (textItems.Length == 0)
        {
            throw new PSArgumentException("Provide at least one Text value in a columnar rich text run.");
        }

        var columns = new Dictionary<string, object?[]>();
        foreach (var propertyNames in SequenceProperties)
        {
            var rawValue = GetValue(value, propertyNames);
            if (rawValue != null)
            {
                columns[propertyNames[0]] = ExpandSequenceValues(rawValue, textItems.Length, propertyNames[0]);
            }
        }

        for (var index = 0; index < textItems.Length; index++)
        {
            var run = new Dictionary<string, object?>
            {
                ["Text"] = textItems[index]
            };

            foreach (var column in columns)
            {
                run[column.Key] = column.Value[index];
            }

            yield return Parse(run);
        }
    }

    private static object?[] ExpandSequenceValues(object value, int textCount, string propertyName)
    {
        value = UnwrapPSObject(value)!;
        if (value is string || value is IDictionary || !(value is IEnumerable values))
        {
            return Enumerable.Repeat<object?>(value, textCount).ToArray();
        }

        var items = values.Cast<object?>().Select(UnwrapPSObject).ToArray();
        if (items.Length == 0)
        {
            throw new PSArgumentException($"Rich text run property '{propertyName}' must contain one value or match the Text count ({textCount}); received 0 values.");
        }

        if (items.Length == 1)
        {
            return Enumerable.Repeat(items[0], textCount).ToArray();
        }

        if (items.Length != textCount)
        {
            throw new PSArgumentException($"Rich text run property '{propertyName}' must contain one value or match the Text count ({textCount}); received {items.Length} values.");
        }

        return items;
    }

    internal static OfficeTextRunSpec Parse(object value)
    {
        if (value is OfficeTextRunSpec spec)
        {
            return NormalizeDerivedFields(spec);
        }

        if (value is string text)
        {
            return new OfficeTextRunSpec { Text = text };
        }

        var kind = GetString(value, "Type", "Kind", "Run");
        var normalizedKind = NormalizeKind(kind);
        var underline = GetUnderline(value, out var underlineStyle) ||
                        normalizedKind is "underline" or "underlined" or "link" or "bookmarklink";

        var baseline = GetString(value, "Baseline");
        if (normalizedKind == "superscript")
        {
            baseline = "Superscript";
        }
        else if (normalizedKind == "subscript")
        {
            baseline = "Subscript";
        }

        return NormalizeDerivedFields(new OfficeTextRunSpec
        {
            Text = GetString(value, "Text", "Value", "Content") ?? string.Empty,
            Kind = kind,
            Bold = GetBool(value, "Bold") || normalizedKind == "bold",
            Italic = GetBool(value, "Italic") || normalizedKind == "italic",
            Underline = underline,
            UnderlineStyle = underlineStyle,
            Strike = GetBool(value, "Strike", "Strikethrough") || normalizedKind is "strike" or "strikethrough",
            Color = GetString(value, "Color", "TextColor", "FontColor"),
            BackgroundColor = GetString(value, "BackgroundColor", "HighlightColor", "FillColor"),
            FontSize = GetDouble(value, "FontSize", "Size"),
            FontName = GetString(value, "FontName", "Font", "Typeface", "FontFamily"),
            Baseline = baseline,
            LinkUri = GetString(value, "LinkUri", "Uri", "Url", "Href"),
            LinkDestinationName = GetString(value, "LinkDestinationName", "DestinationName", "Bookmark", "BookmarkName"),
            LinkContents = GetString(value, "LinkContents", "Contents", "Tooltip"),
            TabLeader = GetString(value, "Leader", "TabLeader"),
            TabAlignment = GetString(value, "Alignment", "TabAlignment")
        });
    }

    internal static string GetPlainText(OfficeTextRunSpec[] runs)
        => string.Concat(runs.Select(run => run.IsLineBreak ? Environment.NewLine : run.IsTab ? "\t" : run.Text));

    internal static string NormalizeKind(string? value)
        => (value ?? string.Empty).Replace("-", string.Empty).Replace("_", string.Empty).Replace(" ", string.Empty).ToLowerInvariant();

    internal static OfficeTextRunSpec NormalizeDerivedFields(OfficeTextRunSpec spec)
    {
        var normalizedKind = NormalizeKind(spec.Kind);
        var baseline = spec.Baseline;
        if (normalizedKind == "superscript")
        {
            baseline = "Superscript";
        }
        else if (normalizedKind == "subscript")
        {
            baseline = "Subscript";
        }

        return new OfficeTextRunSpec
        {
            Text = spec.Text,
            Kind = spec.Kind,
            Bold = spec.Bold || normalizedKind == "bold",
            Italic = spec.Italic || normalizedKind == "italic",
            Underline = spec.Underline || normalizedKind is "underline" or "underlined" or "link" or "bookmarklink",
            UnderlineStyle = spec.UnderlineStyle,
            Strike = spec.Strike || normalizedKind is "strike" or "strikethrough",
            Color = spec.Color,
            BackgroundColor = spec.BackgroundColor,
            FontSize = spec.FontSize,
            FontName = spec.FontName,
            Baseline = baseline,
            LinkUri = spec.LinkUri,
            LinkDestinationName = spec.LinkDestinationName,
            LinkContents = spec.LinkContents,
            TabLeader = spec.TabLeader,
            TabAlignment = spec.TabAlignment
        };
    }

    internal static string? GetString(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value == null ? null : Convert.ToString(UnwrapPSObject(value), CultureInfo.InvariantCulture);
    }

    internal static double? GetDouble(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value == null ? null : Convert.ToDouble(UnwrapPSObject(value), CultureInfo.InvariantCulture);
    }

    internal static bool GetBool(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value != null && Convert.ToBoolean(UnwrapPSObject(value), CultureInfo.InvariantCulture);
    }

    internal static object? GetValue(object? source, params string[] names)
    {
        if (source is PSObject wrapped)
        {
            foreach (var name in names)
            {
                var wrappedProperty = wrapped.Properties
                    .Cast<PSPropertyInfo>()
                    .FirstOrDefault(candidate => candidate.IsGettable && string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase));
                if (wrappedProperty != null)
                {
                    return wrappedProperty.Value;
                }
            }

            source = wrapped.BaseObject;
        }

        if (source is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
                if (names.Any(name => string.Equals(name, key, StringComparison.OrdinalIgnoreCase)))
                {
                    return entry.Value;
                }
            }

            return null;
        }

        if (source is null)
        {
            return null;
        }

        var psObject = PSObject.AsPSObject(source);
        foreach (var name in names)
        {
            var property = psObject.Properties
                .Cast<PSPropertyInfo>()
                .FirstOrDefault(candidate => candidate.IsGettable && string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase));
            if (property != null)
            {
                return property.Value;
            }
        }

        return null;
    }

    private static bool GetUnderline(object source, out string? underlineStyle)
    {
        underlineStyle = null;
        var underlineValue = GetValue(source, "Underline", "Underlined");
        var styleValue = GetValue(source, "UnderlineStyle", "UnderlineKind");
        if (styleValue != null)
        {
            underlineStyle = Convert.ToString(UnwrapPSObject(styleValue), CultureInfo.InvariantCulture);
            return !IsFalseUnderline(underlineStyle);
        }

        if (underlineValue == null)
        {
            return false;
        }

        underlineValue = UnwrapPSObject(underlineValue);
        if (underlineValue is bool boolValue)
        {
            return boolValue;
        }

        underlineStyle = Convert.ToString(underlineValue, CultureInfo.InvariantCulture);
        return !IsFalseUnderline(underlineStyle);
    }

    private static bool IsFalseUnderline(string? value)
    {
        var normalized = NormalizeKind(value);
        return normalized is "" or "false" or "none" or "no" or "off";
    }

    private static object? UnwrapPSObject(object? value)
        => value is PSObject psObject ? psObject.BaseObject : value;
}
