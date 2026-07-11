using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
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

    /// <summary>
    /// Creates a reusable text projector for a fixed CLR object type and column order.
    /// </summary>
#if NET5_0_OR_GREATER
    [System.Diagnostics.CodeAnalysis.RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly or pre-flatten items.")]
#endif
    internal static bool TryCreatePropertyTextProjector(object item, IReadOnlyList<string> columns, out Func<object, string?[], CultureInfo, bool>? projector)
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

        return GetPropertyPlan(item.GetType()).TryCreateTextProjector(columns, out projector);
    }

#if NET5_0_OR_GREATER
    [System.Diagnostics.CodeAnalysis.RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly or pre-flatten items.")]
#endif
    private static ObjectPropertyPlan GetPropertyPlan(Type type) => PropertyPlans.GetOrAdd(type, CreatePropertyPlan);

#if NET5_0_OR_GREATER
    [System.Diagnostics.CodeAnalysis.RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, map values explicitly or pre-flatten items.")]
#endif
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
        private readonly PropertyInfo[] _properties;
        private readonly Dictionary<string, PropertyInfo> _propertiesByName;
        private readonly Type _type;
        private Func<object, object?[], bool>? _defaultProjector;
        private Func<object, string?[], CultureInfo, bool>? _defaultTextProjector;

        public ObjectPropertyPlan(Type type, PropertyInfo[] properties)
        {
            var columnNames = new string[properties.Length];
            _properties = new PropertyInfo[properties.Length];
            _propertiesByName = new Dictionary<string, PropertyInfo>(properties.Length, StringComparer.Ordinal);
            _type = type;
            for (var i = 0; i < properties.Length; i++)
            {
                var property = properties[i];
                _properties[i] = property;
                columnNames[i] = property.Name;
                _propertiesByName[property.Name] = property;
            }

            ColumnNames = columnNames;
        }

        public string[] ColumnNames { get; }

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
            if (IsDefaultColumnOrder(columns))
            {
                projector = _defaultProjector ??= CreateProjector(_properties);
                return true;
            }

            var properties = new PropertyInfo[columns.Count];
            for (var i = 0; i < columns.Count; i++)
            {
                if (!_propertiesByName.TryGetValue(columns[i], out var property))
                {
                    projector = null;
                    return false;
                }

                properties[i] = property;
            }

            projector = CreateProjector(properties);
            return true;
        }

        public bool TryCreateTextProjector(IReadOnlyList<string> columns, out Func<object, string?[], CultureInfo, bool>? projector)
        {
            if (IsDefaultColumnOrder(columns))
            {
                projector = _defaultTextProjector ??= CreateTextProjector(_properties);
                return true;
            }

            var properties = new PropertyInfo[columns.Count];
            for (var i = 0; i < columns.Count; i++)
            {
                if (!_propertiesByName.TryGetValue(columns[i], out var property))
                {
                    projector = null;
                    return false;
                }

                properties[i] = property;
            }

            projector = CreateTextProjector(properties);
            return true;
        }

        private bool IsDefaultColumnOrder(IReadOnlyList<string> columns)
        {
            if (columns.Count != ColumnNames.Length)
            {
                return false;
            }

            for (var i = 0; i < columns.Count; i++)
            {
                if (!string.Equals(columns[i], ColumnNames[i], StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private Func<object, object?[], bool> CreateProjector(PropertyInfo[] properties)
        {
#if NET5_0_OR_GREATER
            if (System.Runtime.CompilerServices.RuntimeFeature.IsDynamicCodeSupported)
            {
                return CreateCompiledProjector(_type, properties);
            }
#endif

            var accessors = new Func<object, object?>[properties.Length];
            for (var i = 0; i < properties.Length; i++)
            {
                accessors[i] = CreateAccessor(properties[i]);
            }

            return CreateProjector(_type, accessors);
        }

        private Func<object, string?[], CultureInfo, bool> CreateTextProjector(PropertyInfo[] properties)
        {
#if NET5_0_OR_GREATER
            if (System.Runtime.CompilerServices.RuntimeFeature.IsDynamicCodeSupported)
            {
                return CreateCompiledTextProjector(_type, properties);
            }
#endif

            var accessors = new Func<object, CultureInfo, string?>[properties.Length];
            for (var i = 0; i < properties.Length; i++)
            {
                accessors[i] = CreateTextAccessor(properties[i]);
            }

            return CreateTextProjector(_type, accessors);
        }

#if NET5_0_OR_GREATER
        private static Func<object, object?[], bool> CreateCompiledProjector(Type type, PropertyInfo[] properties)
        {
            var item = Expression.Parameter(typeof(object), "item");
            var values = Expression.Parameter(typeof(object?[]), "values");
            var typedItem = Expression.Variable(type, "typedItem");
            var returnTarget = Expression.Label(typeof(bool));
            var expressions = new List<Expression>(properties.Length + 6);

            expressions.Add(Expression.IfThen(
                Expression.OrElse(
                    Expression.Equal(item, Expression.Constant(null, typeof(object))),
                    Expression.NotEqual(
                        Expression.Call(item, typeof(object).GetMethod(nameof(GetType))!),
                        Expression.Constant(type))),
                Expression.Return(returnTarget, Expression.Constant(false))));

            expressions.Add(Expression.IfThen(
                Expression.Equal(values, Expression.Constant(null, typeof(object?[]))),
                Expression.Throw(Expression.New(
                    typeof(ArgumentNullException).GetConstructor(new[] { typeof(string) })!,
                    Expression.Constant("values")))));

            expressions.Add(Expression.IfThen(
                Expression.NotEqual(Expression.ArrayLength(values), Expression.Constant(properties.Length)),
                Expression.Throw(Expression.New(
                    typeof(ArgumentException).GetConstructor(new[] { typeof(string), typeof(string) })!,
                    Expression.Constant("Value buffer length must match the projector column count."),
                    Expression.Constant("values")))));

            expressions.Add(Expression.Assign(typedItem, Expression.Convert(item, type)));

            for (var i = 0; i < properties.Length; i++)
            {
                expressions.Add(Expression.Assign(
                    Expression.ArrayAccess(values, Expression.Constant(i)),
                    Expression.Convert(Expression.Property(typedItem, properties[i]), typeof(object))));
            }

            expressions.Add(Expression.Return(returnTarget, Expression.Constant(true)));
            expressions.Add(Expression.Label(returnTarget, Expression.Constant(false)));

            return Expression.Lambda<Func<object, object?[], bool>>(
                Expression.Block(new[] { typedItem }, expressions),
                item,
                values).Compile();
        }

        private static Func<object, string?[], CultureInfo, bool> CreateCompiledTextProjector(Type type, PropertyInfo[] properties)
        {
            var item = Expression.Parameter(typeof(object), "item");
            var values = Expression.Parameter(typeof(string?[]), "values");
            var culture = Expression.Parameter(typeof(CultureInfo), "culture");
            var typedItem = Expression.Variable(type, "typedItem");
            var returnTarget = Expression.Label(typeof(bool));
            var expressions = new List<Expression>(properties.Length + 7);

            expressions.Add(Expression.IfThen(
                Expression.OrElse(
                    Expression.Equal(item, Expression.Constant(null, typeof(object))),
                    Expression.NotEqual(
                        Expression.Call(item, typeof(object).GetMethod(nameof(GetType))!),
                        Expression.Constant(type))),
                Expression.Return(returnTarget, Expression.Constant(false))));

            expressions.Add(Expression.IfThen(
                Expression.Equal(values, Expression.Constant(null, typeof(string?[]))),
                Expression.Throw(Expression.New(
                    typeof(ArgumentNullException).GetConstructor(new[] { typeof(string) })!,
                    Expression.Constant("values")))));

            expressions.Add(Expression.IfThen(
                Expression.Equal(culture, Expression.Constant(null, typeof(CultureInfo))),
                Expression.Throw(Expression.New(
                    typeof(ArgumentNullException).GetConstructor(new[] { typeof(string) })!,
                    Expression.Constant("culture")))));

            expressions.Add(Expression.IfThen(
                Expression.NotEqual(Expression.ArrayLength(values), Expression.Constant(properties.Length)),
                Expression.Throw(Expression.New(
                    typeof(ArgumentException).GetConstructor(new[] { typeof(string), typeof(string) })!,
                    Expression.Constant("Value buffer length must match the projector column count."),
                    Expression.Constant("values")))));

            expressions.Add(Expression.Assign(typedItem, Expression.Convert(item, type)));

            for (var i = 0; i < properties.Length; i++)
            {
                expressions.Add(Expression.Assign(
                    Expression.ArrayAccess(values, Expression.Constant(i)),
                    Expression.Call(
                        typeof(ObjectPropertyPlan).GetMethod(nameof(FormatProjectedValue), BindingFlags.NonPublic | BindingFlags.Static)!,
                        Expression.Convert(Expression.Property(typedItem, properties[i]), typeof(object)),
                        culture)));
            }

            expressions.Add(Expression.Return(returnTarget, Expression.Constant(true)));
            expressions.Add(Expression.Label(returnTarget, Expression.Constant(false)));

            return Expression.Lambda<Func<object, string?[], CultureInfo, bool>>(
                Expression.Block(new[] { typedItem }, expressions),
                item,
                values,
                culture).Compile();
        }
#endif

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

        private static Func<object, string?[], CultureInfo, bool> CreateTextProjector(Type type, Func<object, CultureInfo, string?>[] accessors)
        {
            return (item, values, culture) =>
            {
                if (item == null || item.GetType() != type)
                {
                    return false;
                }

#if NET6_0_OR_GREATER
                ArgumentNullException.ThrowIfNull(values);
                ArgumentNullException.ThrowIfNull(culture);
#else
                if (values == null) throw new ArgumentNullException(nameof(values));
                if (culture == null) throw new ArgumentNullException(nameof(culture));
#endif

                if (values.Length != accessors.Length)
                {
                    throw new ArgumentException("Value buffer length must match the projector column count.", nameof(values));
                }

                for (var i = 0; i < accessors.Length; i++)
                {
                    values[i] = accessors[i](item, culture);
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

        private static Func<object, CultureInfo, string?> CreateTextAccessor(PropertyInfo property)
        {
            var accessor = CreateAccessor(property);
            return (item, culture) => FormatProjectedValue(accessor(item), culture);
        }

        private static string? FormatProjectedValue(object? value, CultureInfo culture)
        {
            if (value is null)
            {
                return null;
            }

            if (value is IFormattable formattable)
            {
                return formattable.ToString(null, culture);
            }

            return value.ToString();
        }
    }
}
