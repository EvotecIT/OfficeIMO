using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Reflection;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Projects typed authoring data once for all three native Office chart writers.</summary>
internal static class OfficeTypedChartData {
    internal static OfficeChartData Create<T, TCategory>(IEnumerable<T> data, Func<T, TCategory> category, OfficeSeries[] series) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (category == null) throw new ArgumentNullException(nameof(category));
        if (series == null) throw new ArgumentNullException(nameof(series));
        var categories = data.Select(row => Convert.ToString(category(row), CultureInfo.InvariantCulture) ?? string.Empty).ToArray();
        foreach (var item in series) {
            if (item == null) throw new ArgumentException("Series must not contain null entries.", nameof(series));
            if (item.Values.Count != categories.Length) throw new ArgumentException("Each chart series must have one value per category.", nameof(series));
        }
        return new OfficeChartData(categories, series);
    }

    internal static OfficeChartData Create<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        IEnumerable<T> data, string category, string[] values) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (values == null) throw new ArgumentNullException(nameof(values));
        if (values.Length == 0) throw new ArgumentException("Select at least one numeric series property.", nameof(values));
        var categoryProperty = Property(typeof(T), category);
        var numericProperties = values.Select(name => Property(typeof(T), name)).ToArray();
        foreach (var property in numericProperties)
            if (!OfficeSeries.IsNumeric(property.PropertyType)) throw new ArgumentException($"Series property '{property.Name}' must be numeric.", nameof(values));
        var labels = new List<string>();
        var columns = numericProperties.Select(_ => new List<double>()).ToArray();
        foreach (T row in data) {
            if (row is null) throw new ArgumentException("Chart rows must not contain null entries.", nameof(data));
            labels.Add(Convert.ToString(categoryProperty.GetValue(row, null), CultureInfo.InvariantCulture) ?? string.Empty);
            for (int index = 0; index < columns.Length; index++) columns[index].Add(OfficeSeries.NumericValue(numericProperties[index].GetValue(row, null)));
        }
        return new OfficeChartData(labels, numericProperties.Select((property, index) => new OfficeSeries(property.Name, columns[index])));
    }

    private static PropertyInfo Property([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] Type type, string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Supply a readable property name.", nameof(name));
        var property = type.GetProperty(name, BindingFlags.Instance | BindingFlags.Public);
        if (property?.GetMethod == null || property.GetIndexParameters().Length != 0)
            throw new ArgumentException($"'{name}' is not a readable scalar property of {type.Name}.", nameof(name));
        return property;
    }
}
