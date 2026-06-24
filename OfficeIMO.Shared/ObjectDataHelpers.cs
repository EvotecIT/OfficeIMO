using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace OfficeIMO.Shared;

/// <summary>
/// Shared helpers for projecting object data into tabular structures.
/// </summary>
internal static class ObjectDataHelpers
{
    private static readonly ConcurrentDictionary<Type, ObjectPropertyPlan> PropertyPlans = new();

    /// <summary>
    /// Resolves column names from a dictionary or public readable properties.
    /// </summary>
#if NET5_0_OR_GREATER
    [System.Diagnostics.CodeAnalysis.RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly or pre-flatten items.")]
#endif
    public static IReadOnlyList<string> GetColumnNames(object item)
    {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(item);
#else
        if (item == null) throw new ArgumentNullException(nameof(item));
#endif

        if (item is Dictionary<string, object?> dictionary)
        {
            return dictionary.Keys.Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
        }

        if (item is IReadOnlyDictionary<string, object?> roDict)
        {
            return roDict.Keys.Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
        }

        if (item is IDictionary<string, object?> dict)
        {
            return dict.Keys.Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
        }

        if (item is IDictionary legacyDict)
        {
            var names = new List<string>();
            foreach (DictionaryEntry entry in legacyDict)
            {
                var key = entry.Key?.ToString();
                if (!string.IsNullOrWhiteSpace(key))
                {
                    names.Add(key!);
                }
            }
            return names;
        }

        return GetPropertyPlan(item.GetType()).ColumnNames;
    }

    /// <summary>
    /// Retrieves a value for the specified column from a dictionary or property.
    /// </summary>
#if NET5_0_OR_GREATER
    [System.Diagnostics.CodeAnalysis.RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly or pre-flatten items.")]
#endif
    public static object? GetValue(object item, string column)
    {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(item);
        ArgumentNullException.ThrowIfNull(column);
#else
        if (item == null) throw new ArgumentNullException(nameof(item));
        if (column == null) throw new ArgumentNullException(nameof(column));
#endif

        if (item is Dictionary<string, object?> dictionary)
        {
            return dictionary.TryGetValue(column, out var value) ? value : null;
        }

        if (item is IReadOnlyDictionary<string, object?> roDict)
        {
            return roDict.TryGetValue(column, out var value) ? value : null;
        }

        if (item is IDictionary<string, object?> dict)
        {
            return dict.TryGetValue(column, out var value) ? value : null;
        }

        if (item is IDictionary legacyDict)
        {
            if (legacyDict.Contains(column))
            {
                return legacyDict[column];
            }

            foreach (DictionaryEntry entry in legacyDict)
            {
                var key = entry.Key?.ToString();
                if (string.Equals(key, column, StringComparison.OrdinalIgnoreCase))
                {
                    return entry.Value;
                }
            }

            return null;
        }

        return GetPropertyPlan(item.GetType()).TryGetValue(item, column, out var propertyValue) ? propertyValue : null;
    }

    private static ObjectPropertyPlan GetPropertyPlan(Type type) => PropertyPlans.GetOrAdd(type, CreatePropertyPlan);

    private static ObjectPropertyPlan CreatePropertyPlan(Type type)
    {
        var properties = type
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(static p => p.CanRead && p.GetIndexParameters().Length == 0 && !string.IsNullOrWhiteSpace(p.Name))
            .OrderBy(static p => p.MetadataToken)
            .ToArray();

        return new ObjectPropertyPlan(properties);
    }

    private sealed class ObjectPropertyPlan
    {
        private readonly Dictionary<string, PropertyInfo> _propertiesByName;

        public ObjectPropertyPlan(IReadOnlyList<PropertyInfo> properties)
        {
            var columnNames = new string[properties.Count];
            _propertiesByName = new Dictionary<string, PropertyInfo>(properties.Count, StringComparer.Ordinal);
            for (var i = 0; i < properties.Count; i++)
            {
                var property = properties[i];
                columnNames[i] = property.Name;
                _propertiesByName[property.Name] = property;
            }

            ColumnNames = columnNames;
        }

        public IReadOnlyList<string> ColumnNames { get; }

        public bool TryGetValue(object item, string column, out object? value)
        {
            if (_propertiesByName.TryGetValue(column, out var property))
            {
                value = property.GetValue(item);
                return true;
            }

            value = null;
            return false;
        }
    }
}
