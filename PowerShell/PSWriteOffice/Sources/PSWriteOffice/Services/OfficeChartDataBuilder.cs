using System;
using System.Collections;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;

namespace PSWriteOffice.Services;

/// <summary>Projects PowerShell property and dictionary values into the existing native chart data contract.</summary>
internal static class OfficeChartDataBuilder {
    internal static OfficeChartData Build(object[] items, string category, string[] values, OfficeChartType type) {
        if (items.Length == 0) throw new PSArgumentException("Provide at least one chart data item.", nameof(items));
        if (values.Length == 0) throw new PSArgumentException("Provide at least one Value property.", nameof(values));
        var x = type == OfficeChartType.Scatter ? items.Select(item => ConvertToDouble(GetPropertyValue(item, category), category)).ToArray() : null;
        var categories = x != null ? x.Select(value => value.ToString(CultureInfo.InvariantCulture)).ToArray()
            : items.Select(item => ConvertToString(GetPropertyValue(item, category), category)).ToArray();
        var series = values.Select(property => new OfficeSeries(property, items.Select(item => ConvertToDouble(GetPropertyValue(item, property), property)), x)).ToArray();
        return new OfficeChartData(categories, series);
    }
    internal static object? GetPropertyValue(object item, string propertyName) {
        if (item == null) {
            throw new PSArgumentException("Chart data items cannot be null.");
        }

        if (string.IsNullOrWhiteSpace(propertyName)) {
            throw new PSArgumentException("Property name cannot be empty.", nameof(propertyName));
        }

        if (item is IDictionary dictionary) {
            foreach (DictionaryEntry entry in dictionary) {
                if (entry.Key is string key && string.Equals(key, propertyName, StringComparison.OrdinalIgnoreCase)) {
                    return UnwrapValue(entry.Value);
                }
            }
        }

        var property = PSObject.AsPSObject(item).Properties[propertyName];
        if (property == null) {
            throw new PSArgumentException($"Property '{propertyName}' was not found on chart data item.");
        }

        return UnwrapValue(property.Value);
    }

    internal static object? UnwrapValue(object? value) {
        // Select-Object calculated properties can retain a PSObject around a scalar.
        while (value is PSObject wrapper && !ReferenceEquals(value, wrapper.BaseObject)) {
            value = wrapper.BaseObject;
        }

        return value;
    }

    internal static string ConvertToString(object? value, string propertyName) {
        if (value == null) {
            throw new PSArgumentException($"Property '{propertyName}' cannot be null.");
        }

        var text = Convert.ToString(value, CultureInfo.InvariantCulture);
        if (string.IsNullOrWhiteSpace(text)) {
            throw new PSArgumentException($"Property '{propertyName}' cannot be empty.");
        }

        return text;
    }

    internal static double ConvertToDouble(object? value, string propertyName) {
        if (value == null) {
            throw new PSArgumentException($"Property '{propertyName}' cannot be null.");
        }

        try {
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        } catch (Exception error) when (error is FormatException or InvalidCastException or OverflowException) {
            throw new PSArgumentException($"Property '{propertyName}' must be numeric.", propertyName);
        }
    }
}
