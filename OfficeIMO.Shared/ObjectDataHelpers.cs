using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace OfficeIMO.Shared;

/// <summary>
/// Shared helpers for projecting object data into tabular structures.
/// </summary>
internal static class ObjectDataHelpers
{
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

        var props = item.GetType()
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead && p.GetIndexParameters().Length == 0)
            .OrderBy(p => p.MetadataToken)
            .Select(p => p.Name)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .ToList();

        return props;
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

        var prop = item.GetType().GetProperty(column, BindingFlags.Public | BindingFlags.Instance);
        return prop?.GetValue(item);
    }
}
