using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Services.PowerPoint;

internal static class PowerPointDesignerDataMapper
{
    public static IReadOnlyList<PowerPointProcessStep> ToProcessSteps(object[] data)
    {
        return EnsureData(data, "Process steps")
            .Select(item => new PowerPointProcessStep(
                GetRequiredString(item, "Title"),
                GetRequiredString(item, "Body", "Description", "Text"),
                GetOptionalString(item, "Number")))
            .ToList();
    }

    public static IReadOnlyList<PowerPointCardContent> ToCards(object[] data)
    {
        return EnsureData(data, "Cards")
            .Select(item => new PowerPointCardContent(
                GetRequiredString(item, "Title", "Name"),
                GetStringList(item, "Items", "Bullets", "Details"),
                GetOptionalColor(item, "AccentColor", "Color")))
            .ToList();
    }

    public static IReadOnlyList<PowerPointCoverageLocation> ToLocations(object[] data)
    {
        return EnsureData(data, "Coverage locations")
            .Select(item => new PowerPointCoverageLocation(
                GetRequiredString(item, "Name", "Title"),
                GetRequiredDouble(item, "X"),
                GetRequiredDouble(item, "Y"),
                GetOptionalString(item, "Detail", "Subtitle", "Description")))
            .ToList();
    }

    public static IReadOnlyList<PowerPointCapabilitySection> ToCapabilitySections(object[] data)
    {
        return EnsureData(data, "Capability sections")
            .Select(item => new PowerPointCapabilitySection(
                GetRequiredString(item, "Heading", "Title", "Name"),
                GetOptionalString(item, "Body", "Description", "Text"),
                GetStringList(item, "Items", "Bullets", "Details"),
                GetOptionalColor(item, "AccentColor", "Color")))
            .ToList();
    }

    public static IReadOnlyList<PowerPointLogoItem> ToLogoItems(object[] data)
    {
        return EnsureData(data, "Logo items")
            .Select(item => new PowerPointLogoItem(
                GetRequiredString(item, "Name", "Title"),
                GetOptionalString(item, "Subtitle", "Detail", "Description"),
                GetOptionalString(item, "ImagePath", "Path"),
                GetOptionalColor(item, "AccentColor", "Color")))
            .ToList();
    }

    public static IReadOnlyList<PowerPointCaseStudySection> ToCaseStudySections(object[] data)
    {
        return EnsureData(data, "Case-study sections")
            .Select(item => new PowerPointCaseStudySection(
                GetRequiredString(item, "Heading", "Title", "Name"),
                GetRequiredString(item, "Body", "Description", "Text")))
            .ToList();
    }

    public static IReadOnlyList<PowerPointMetric> ToMetrics(object[]? data)
    {
        if (data == null || data.Length == 0)
        {
            return Array.Empty<PowerPointMetric>();
        }

        return EnsureData(data, "Metrics")
            .Select(item => new PowerPointMetric(
                GetRequiredString(item, "Value"),
                GetRequiredString(item, "Label", "Name", "Title")))
            .ToList();
    }

    private static IEnumerable<object> EnsureData(object[]? data, string name)
    {
        var items = data?.Where(item => item != null).ToArray();
        if (items == null || items.Length == 0)
        {
            throw new PSArgumentException($"{name} require at least one item.");
        }

        return items;
    }

    private static string GetRequiredString(object item, params string[] names)
    {
        var value = GetOptionalString(item, names);
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new PSArgumentException($"Input item is missing required property '{names[0]}'.");
        }

        return value!;
    }

    private static string? GetOptionalString(object item, params string[] names)
    {
        var value = GetPropertyValue(item, names);
        return value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    private static string? GetOptionalColor(object item, params string[] names)
    {
        var value = GetOptionalString(item, names);
        return string.IsNullOrWhiteSpace(value) ? null : value!.Trim().TrimStart('#');
    }

    private static double GetRequiredDouble(object item, string name)
    {
        var value = GetPropertyValue(item, name);
        if (value == null)
        {
            throw new PSArgumentException($"Input item is missing required property '{name}'.");
        }

        try
        {
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }
        catch (Exception)
        {
            throw new PSArgumentException($"Property '{name}' must be numeric.");
        }
    }

    private static IReadOnlyList<string> GetStringList(object item, params string[] names)
    {
        var value = GetPropertyValue(item, names);
        if (value == null)
        {
            return Array.Empty<string>();
        }

        if (value is string text)
        {
            return SplitText(text);
        }

        if (value is IEnumerable enumerable)
        {
            var result = new List<string>();
            foreach (var entry in enumerable)
            {
                if (entry == null)
                {
                    continue;
                }

                var textValue = Convert.ToString(entry, CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(textValue))
                {
                    result.Add(textValue!);
                }
            }

            return result;
        }

        var single = Convert.ToString(value, CultureInfo.InvariantCulture);
        return string.IsNullOrWhiteSpace(single)
            ? Array.Empty<string>()
            : new[] { single! };
    }

    private static IReadOnlyList<string> SplitText(string text)
    {
        return text
            .Split(new[] { '|', ';' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(item => item.Trim())
            .Where(item => item.Length > 0)
            .ToList();
    }

    private static object? GetPropertyValue(object item, params string[] names)
    {
        if (item is IDictionary dictionary)
        {
            foreach (var name in names)
            {
                foreach (DictionaryEntry entry in dictionary)
                {
                    if (entry.Key is string key && string.Equals(key, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return entry.Value;
                    }
                }
            }
        }

        var psObject = PSObject.AsPSObject(item);
        foreach (var name in names)
        {
            var property = psObject.Properties[name];
            if (property != null)
            {
                return property.Value;
            }
        }

        return null;
    }
}
