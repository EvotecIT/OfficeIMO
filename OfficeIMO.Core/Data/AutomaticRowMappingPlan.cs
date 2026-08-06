#nullable enable

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization;
#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
#endif

namespace OfficeIMO.Data;

internal sealed class AutomaticRowMappingPlan<
    [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T> where T : new() {
    private readonly MappingBinding[] _bindings;

    private AutomaticRowMappingPlan(MappingBinding[] bindings) {
        _bindings = bindings;
    }

    internal static AutomaticRowMappingPlan<T> Create(
        IReadOnlyList<string> headers,
        bool requireAllColumnsMapped = false) {
        AutomaticPropertySetter?[] mapped = new AutomaticPropertySetter?[headers.Count];
        var assigned = new HashSet<AutomaticPropertySetter>();
        MapPass(headers, mapped, assigned, AutomaticMappingCache.Exact, static value => value, "with that exact name");
        MapPass(headers, mapped, assigned, AutomaticMappingCache.Insensitive, static value => value, "when casing is ignored");
        MapPass(headers, mapped, assigned, AutomaticMappingCache.Aliases, static value => value, "through a declared alias");
        MapPass(headers, mapped, assigned, AutomaticMappingCache.Canonical, Canonicalize, "when punctuation is ignored");
        MapPass(headers, mapped, assigned, AutomaticMappingCache.CanonicalAliases, Canonicalize, "through a normalized alias");

        if (requireAllColumnsMapped) {
            string[] unmappedHeaders = headers
                .Where((header, index) => mapped[index] is null && !string.IsNullOrWhiteSpace(header))
                .ToArray();
            if (unmappedHeaders.Length > 0) {
                throw new DataMappingException(
                    $"Typed mapping for '{typeof(T).Name}' is strict and could not resolve columns: {string.Join(", ", unmappedHeaders.Select(static header => $"'{header}'"))}.");
            }
        }

        MappingBinding[] bindings = mapped
            .Select((property, index) => property is null
                ? default(MappingBinding?)
                : new MappingBinding(index, property))
            .Where(static binding => binding.HasValue)
            .Select(static binding => binding!.Value)
            .ToArray();

        if (bindings.Length == 0) {
            throw new DataMappingException($"No columns match writable properties on {typeof(T).Name}.");
        }
        return new AutomaticRowMappingPlan<T>(bindings);
    }

    internal T MapRow(
        Func<int, object?> getValue,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter = null) {
        T instance = new T();
        foreach (MappingBinding binding in _bindings) {
            instance = binding.Apply(instance, getValue(binding.ColumnIndex), culture, dateTimeFormats, typeConverter);
        }
        return instance;
    }

    private static void MapPass(
        IReadOnlyList<string> headers,
        AutomaticPropertySetter?[] mapped,
        HashSet<AutomaticPropertySetter> assigned,
        IReadOnlyDictionary<string, AutomaticPropertySetter[]> lookup,
        Func<string, string> keySelector,
        string ambiguityDescription) {
        for (int columnIndex = 0; columnIndex < headers.Count; columnIndex++) {
            if (mapped[columnIndex] is not null) continue;
            string key = keySelector(headers[columnIndex]);
            if (key.Length == 0 || !lookup.TryGetValue(key, out AutomaticPropertySetter[]? candidates)) continue;

            AutomaticPropertySetter? match = null;
            foreach (AutomaticPropertySetter candidate in candidates) {
                if (assigned.Contains(candidate)) continue;
                if (match is not null) {
                    throw new DataMappingException(
                        $"Column '{headers[columnIndex]}' matches multiple writable properties on {typeof(T).Name} {ambiguityDescription}.");
                }
                match = candidate;
            }
            if (match is null) continue;
            mapped[columnIndex] = match;
            assigned.Add(match);
        }
    }

    private static string Canonicalize(string value) {
        var builder = new System.Text.StringBuilder(value.Length);
        foreach (char character in value) {
            if (char.IsLetterOrDigit(character)) builder.Append(char.ToUpperInvariant(character));
        }
        return builder.ToString();
    }

    private readonly struct MappingBinding {
        private readonly AutomaticPropertySetter _property;

        internal MappingBinding(int columnIndex, AutomaticPropertySetter property) {
            ColumnIndex = columnIndex;
            _property = property;
        }

        internal int ColumnIndex { get; }

        internal T Apply(
            T instance,
            object? rawValue,
            CultureInfo culture,
            IReadOnlyList<string>? dateTimeFormats,
            Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter) {
            if (!DataValueConverter.TryConvert(rawValue, _property.ValueType, culture, dateTimeFormats, typeConverter, out object? converted, out string? error)) {
                throw new DataMappingException(
                    $"Column value cannot be assigned to {typeof(T).Name}.{_property.Name}: {error}");
            }
            return _property.Assign(instance, converted);
        }
    }

    private sealed class AutomaticPropertySetter {
        internal AutomaticPropertySetter(PropertyInfo property, Func<T, object?, T> assign) {
            Property = property;
            Name = property.Name;
            ValueType = property.PropertyType;
            Assign = assign;
        }

        internal PropertyInfo Property { get; }
        internal string Name { get; }
        internal Type ValueType { get; }
        internal Func<T, object?, T> Assign { get; }
    }

    private static class AutomaticMappingCache {
        internal static readonly Dictionary<string, AutomaticPropertySetter[]> Exact;
        internal static readonly Dictionary<string, AutomaticPropertySetter[]> Insensitive;
        internal static readonly Dictionary<string, AutomaticPropertySetter[]> Aliases;
        internal static readonly Dictionary<string, AutomaticPropertySetter[]> Canonical;
        internal static readonly Dictionary<string, AutomaticPropertySetter[]> CanonicalAliases;

        static AutomaticMappingCache() {
            AutomaticPropertySetter[] properties = typeof(T)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(static property => property.SetMethod?.IsPublic == true && property.GetIndexParameters().Length == 0)
                .Select(static property => new AutomaticPropertySetter(property, CreateAssignment(property)))
                .ToArray();
            Exact = CreateLookup(properties, static property => property.Name, StringComparer.Ordinal);
            Insensitive = CreateLookup(properties, static property => property.Name, StringComparer.OrdinalIgnoreCase);
            Aliases = CreateAliasLookup(properties, canonical: false);
            Canonical = CreateLookup(properties, static property => Canonicalize(property.Name), StringComparer.Ordinal);
            CanonicalAliases = CreateAliasLookup(properties, canonical: true);
        }
    }

    private static Dictionary<string, AutomaticPropertySetter[]> CreateAliasLookup(
        IEnumerable<AutomaticPropertySetter> properties,
        bool canonical) {
        IEqualityComparer<string> comparer = canonical ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
        return properties
            .SelectMany(property => GetAliases(property.Property)
                .Select(alias => new { Key = canonical ? Canonicalize(alias) : alias, Property = property }))
            .Where(static item => item.Key.Length > 0)
            .GroupBy(static item => item.Key, comparer)
            .ToDictionary(
                static group => group.Key,
                static group => group.Select(static item => item.Property).Distinct().ToArray(),
                comparer);
    }

    private static IEnumerable<string> GetAliases(PropertyInfo property) {
        var aliases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (property.GetCustomAttribute<DisplayNameAttribute>(inherit: true)?.DisplayName is { } displayName &&
            !string.IsNullOrWhiteSpace(displayName)) {
            aliases.Add(displayName);
        }
        if (property.GetCustomAttribute<DataMemberAttribute>(inherit: true)?.Name is { } dataMemberName &&
            !string.IsNullOrWhiteSpace(dataMemberName)) {
            aliases.Add(dataMemberName);
        }
        foreach (IDataColumnAliasProvider provider in property
                     .GetCustomAttributes(inherit: true)
                     .OfType<IDataColumnAliasProvider>()) {
            foreach (string alias in provider.ColumnAliases) {
                if (!string.IsNullOrWhiteSpace(alias)) aliases.Add(alias);
            }
        }
        return aliases;
    }

    private static Dictionary<string, AutomaticPropertySetter[]> CreateLookup(
        IEnumerable<AutomaticPropertySetter> properties,
        Func<AutomaticPropertySetter, string> getKey,
        IEqualityComparer<string> comparer) {
        return properties
            .Select(property => new { Key = getKey(property), Property = property })
            .Where(static item => item.Key.Length > 0)
            .GroupBy(static item => item.Key, comparer)
            .ToDictionary(
                static group => group.Key,
                static group => group.Select(static item => item.Property).ToArray(),
                comparer);
    }

    private static Func<T, object?, T> CreateAssignment(PropertyInfo property) {
#if NET8_0_OR_GREATER
        if (!RuntimeFeature.IsDynamicCodeSupported) return CreateReflectionAssignment(property);
#endif
        try {
            ParameterExpression target = Expression.Parameter(typeof(T), "target");
            ParameterExpression value = Expression.Parameter(typeof(object), "value");
            BinaryExpression assignment = Expression.Assign(
                Expression.Property(target, property),
                Expression.Convert(value, property.PropertyType));
            BlockExpression body = Expression.Block(assignment, target);
            return Expression.Lambda<Func<T, object?, T>>(body, target, value).Compile();
        } catch {
            return CreateReflectionAssignment(property);
        }
    }

    private static Func<T, object?, T> CreateReflectionAssignment(PropertyInfo property) =>
        (target, value) => {
            object boxed = target!;
            property.SetValue(boxed, value, index: null);
            return (T)boxed;
        };
}
