using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
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
    /// Returns whether the item exposes dictionary keys rather than fixed CLR properties.
    /// </summary>
    public static bool IsDictionaryLike(object item)
    {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(item);
#else
        if (item == null) throw new ArgumentNullException(nameof(item));
#endif

        return item is Dictionary<string, object?>
            || item is IReadOnlyDictionary<string, object?>
            || item is IDictionary<string, object?>
            || item is IDictionary;
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

    /// <summary>
    /// Creates a reusable projector for a fixed CLR object type and column order.
    /// </summary>
#if NET5_0_OR_GREATER
    [System.Diagnostics.CodeAnalysis.RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly or pre-flatten items.")]
#endif
    internal static bool TryCreatePropertyProjector(object item, IReadOnlyList<string> columns, out Func<object, object?[], bool>? projector)
    {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(item);
        ArgumentNullException.ThrowIfNull(columns);
#else
        if (item == null) throw new ArgumentNullException(nameof(item));
        if (columns == null) throw new ArgumentNullException(nameof(columns));
#endif

        projector = null;
        if (IsDictionaryLike(item))
        {
            return false;
        }

        return GetPropertyPlan(item.GetType()).TryCreateProjector(columns, out projector);
    }

    private static ObjectPropertyPlan GetPropertyPlan(Type type) => PropertyPlans.GetOrAdd(type, CreatePropertyPlan);

    private static ObjectPropertyPlan CreatePropertyPlan(Type type)
    {
        var properties = type
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(static p => p.CanRead && p.GetIndexParameters().Length == 0 && !string.IsNullOrWhiteSpace(p.Name))
            .OrderBy(static p => p.MetadataToken)
            .ToArray();

        return new ObjectPropertyPlan(type, properties);
    }

    private sealed class ObjectPropertyPlan
    {
        private readonly Dictionary<string, PropertyInfo> _propertiesByName;
        private readonly Type _type;

        public ObjectPropertyPlan(Type type, IReadOnlyList<PropertyInfo> properties)
        {
            var columnNames = new string[properties.Count];
            _propertiesByName = new Dictionary<string, PropertyInfo>(properties.Count, StringComparer.Ordinal);
            _type = type;
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

        public bool TryCreateProjector(IReadOnlyList<string> columns, out Func<object, object?[], bool>? projector)
        {
            var accessors = new Func<object, object?>[columns.Count];
            for (var i = 0; i < columns.Count; i++)
            {
                if (!_propertiesByName.TryGetValue(columns[i], out var property))
                {
                    projector = null;
                    return false;
                }

                accessors[i] = CreateAccessor(property);
            }

            projector = CreateProjector(_type, accessors);
            return true;
        }

        private static Func<object, object?[], bool> CreateProjector(Type type, Func<object, object?>[] accessors)
        {
            return (item, values) =>
            {
                if (item == null || item.GetType() != type)
                {
                    return false;
                }

#if NET6_0_OR_GREATER
                ArgumentNullException.ThrowIfNull(values);
#else
                if (values == null) throw new ArgumentNullException(nameof(values));
#endif

                if (values.Length != accessors.Length)
                {
                    throw new ArgumentException("Value buffer length must match the projector column count.", nameof(values));
                }

                for (var i = 0; i < accessors.Length; i++)
                {
                    values[i] = accessors[i](item);
                }

                return true;
            };
        }

        private static Func<object, object?> CreateAccessor(PropertyInfo property)
        {
#if NET5_0_OR_GREATER
            if (System.Runtime.CompilerServices.RuntimeFeature.IsDynamicCodeSupported)
            {
                var item = Expression.Parameter(typeof(object), "item");
                var typedItem = Expression.Convert(item, property.DeclaringType!);
                var value = Expression.Property(typedItem, property);
                var boxedValue = Expression.Convert(value, typeof(object));
                return Expression.Lambda<Func<object, object?>>(boxedValue, item).Compile();
            }
#endif

            return property.GetValue;
        }
    }
}
