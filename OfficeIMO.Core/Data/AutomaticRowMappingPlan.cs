#nullable enable

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Common;
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
    private readonly Func<DbDataReader, T>? _fastReaderMap;
    private static CacheEntry? _cachedPlan;

    private AutomaticRowMappingPlan(
        MappingBinding[] bindings,
        Func<DbDataReader, T>? fastReaderMap) {
        _bindings = bindings;
        _fastReaderMap = fastReaderMap;
    }

    internal static AutomaticRowMappingPlan<T> Create(
        IReadOnlyList<string> headers,
        bool requireAllColumnsMapped = false) {
        CacheEntry? cached = System.Threading.Volatile.Read(ref _cachedPlan);
        if (cached is not null && cached.Matches(headers, requireAllColumnsMapped)) {
            return cached.Plan;
        }

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
        var plan = new AutomaticRowMappingPlan<T>(bindings, CreateFastReaderMap(bindings));
        System.Threading.Volatile.Write(
            ref _cachedPlan,
            new CacheEntry(headers.ToArray(), requireAllColumnsMapped, plan));
        return plan;
    }

    internal T MapRow(
        Func<int, object?> getValue,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter = null,
        DataMappingErrorValuePolicy errorValuePolicy = DataMappingErrorValuePolicy.Include) {
        T instance = new T();
        foreach (MappingBinding binding in _bindings) {
            instance = binding.Apply(instance, getValue(binding.ColumnIndex), culture, dateTimeFormats, typeConverter, errorValuePolicy);
        }
        return instance;
    }

    internal T MapValues(
        object?[] values,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter = null,
        DataMappingErrorValuePolicy errorValuePolicy = DataMappingErrorValuePolicy.Include) {
        T instance = new T();
        foreach (MappingBinding binding in _bindings) {
            instance = binding.Apply(
                instance,
                values[binding.ColumnIndex],
                culture,
                dateTimeFormats,
                typeConverter,
                errorValuePolicy);
        }
        return instance;
    }

    internal object?[] CaptureReaderValues(DbDataReader reader) {
        var values = new object?[reader.FieldCount];
        for (int index = 0; index < values.Length; index++) {
            values[index] = DBNull.Value;
        }
        foreach (MappingBinding binding in _bindings) {
            values[binding.ColumnIndex] = binding.Property.ReadReaderValue(reader, binding.ColumnIndex);
        }
        return values;
    }

    private bool HasNullReaderValue(DbDataReader reader) {
        foreach (MappingBinding binding in _bindings) {
            if (reader.IsDBNull(binding.ColumnIndex)) return true;
        }
        return false;
    }

    internal T MapReaderRow(
        DbDataReader reader,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter = null,
        DataMappingErrorValuePolicy errorValuePolicy = DataMappingErrorValuePolicy.Include) {
        if (typeConverter is null &&
            _fastReaderMap is not null &&
            reader is IDataReaderFastMappingValues &&
            !HasNullReaderValue(reader)) {
            try {
                return _fastReaderMap(reader);
            } catch (TypedReaderValueException) {
            }
        }

        T instance = new T();
        foreach (MappingBinding binding in _bindings) {
            instance = binding.ApplyReader(
                instance,
                reader,
                culture,
                dateTimeFormats,
                typeConverter,
                errorValuePolicy);
        }
        return instance;
    }

    internal Func<DbDataReader, T>? GetFastReaderMap(
        DbDataReader reader,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter) =>
        typeConverter is null &&
        _fastReaderMap is not null &&
        reader is IDataReaderFastMappingValues { HasOnlyNonNullFastValues: true }
            ? _fastReaderMap
            : null;

    internal T MapReaderRowSlow(
        DbDataReader reader,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
        DataMappingErrorValuePolicy errorValuePolicy) {
        T instance = new T();
        foreach (MappingBinding binding in _bindings) {
            instance = binding.ApplyReader(
                instance,
                reader,
                culture,
                dateTimeFormats,
                typeConverter,
                errorValuePolicy);
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

        internal AutomaticPropertySetter Property => _property;

        internal T Apply(
            T instance,
            object? rawValue,
            CultureInfo culture,
            IReadOnlyList<string>? dateTimeFormats,
            Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
            DataMappingErrorValuePolicy errorValuePolicy) {
            if (!DataValueConverter.TryConvert(rawValue, _property.ValueType, culture, dateTimeFormats, typeConverter, errorValuePolicy, out object? converted, out string? error)) {
                throw new DataMappingException(
                    $"Column value cannot be assigned to {typeof(T).Name}.{_property.Name}: {error}");
            }
            return _property.Assign(instance, converted);
        }

        internal T ApplyReader(
            T instance,
            DbDataReader reader,
            CultureInfo culture,
            IReadOnlyList<string>? dateTimeFormats,
            Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
            DataMappingErrorValuePolicy errorValuePolicy) {
            return _property.ApplyReader(
                instance,
                reader,
                ColumnIndex,
                culture,
                dateTimeFormats,
                typeConverter,
                errorValuePolicy);
        }
    }

    private sealed class AutomaticPropertySetter {
        private readonly Func<T, object?, T> _assign;
        private readonly Func<T, DbDataReader, int, T>? _assignReader;
        private readonly Func<DbDataReader, int, object?>? _readReaderValue;

        internal AutomaticPropertySetter(
            PropertyInfo property,
            Func<T, object?, T> assign,
            Func<T, DbDataReader, int, T>? assignReader,
            Func<DbDataReader, int, object?>? readReaderValue) {
            Property = property;
            Name = property.Name;
            ValueType = property.PropertyType;
            _assign = assign;
            _assignReader = assignReader;
            _readReaderValue = readReaderValue;
        }

        internal PropertyInfo Property { get; }
        internal string Name { get; }
        internal Type ValueType { get; }

        internal T Assign(T instance, object? value) => _assign(instance, value);

        internal object? ReadReaderValue(DbDataReader reader, int ordinal) {
            if (reader.IsDBNull(ordinal)) return null;
            if (_readReaderValue is not null) {
                try {
                    return _readReaderValue(reader, ordinal);
                } catch (InvalidCastException) {
                } catch (FormatException) {
                } catch (OverflowException) {
                } catch (NotSupportedException) {
                } catch (NotImplementedException) {
                }
            }
            object value = reader.GetValue(ordinal);
            return ReferenceEquals(value, DBNull.Value) ? null : value;
        }

        internal T ApplyReader(
            T instance,
            DbDataReader reader,
            int ordinal,
            CultureInfo culture,
            IReadOnlyList<string>? dateTimeFormats,
            Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
            DataMappingErrorValuePolicy errorValuePolicy) {
            if (typeConverter is null &&
                reader is IDataReaderFastMappingValues &&
                !reader.IsDBNull(ordinal) &&
                _assignReader is not null) {
                try {
                    return _assignReader(instance, reader, ordinal);
                } catch (TypedReaderValueException) {
                }
            }

            object? rawValue = reader.GetValue(ordinal);
            if (!DataValueConverter.TryConvert(
                ReferenceEquals(rawValue, DBNull.Value) ? null : rawValue,
                ValueType,
                culture,
                dateTimeFormats,
                typeConverter,
                errorValuePolicy,
                out object? converted,
                out string? error)) {
                throw new DataMappingException(
                    $"Column value cannot be assigned to {typeof(T).Name}.{Name}: {error}");
            }
            return _assign(instance, converted);
        }
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
                .Select(static property => new AutomaticPropertySetter(
                    property,
                    CreateAssignment(property),
                    CreateReaderAssignment(property),
                    CreateReaderValueAccessor(property)))
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

    private static Func<T, DbDataReader, int, T>? CreateReaderAssignment(PropertyInfo property) {
#if NET8_0_OR_GREATER
        if (!RuntimeFeature.IsDynamicCodeSupported) return null;
#endif
        MethodInfo? getter = GetTypedGetter(property.PropertyType);
        if (getter is null) return null;

        try {
            ParameterExpression target = Expression.Parameter(typeof(T), "target");
            ParameterExpression reader = Expression.Parameter(typeof(DbDataReader), "reader");
            ParameterExpression ordinal = Expression.Parameter(typeof(int), "ordinal");
            ParameterExpression value = Expression.Variable(property.PropertyType, "value");
            NewExpression fallbackException = Expression.New(typeof(TypedReaderValueException));
            BinaryExpression readValue = Expression.Assign(
                value,
                Expression.Call(reader, getter, ordinal));
            TryExpression guardedRead = Expression.TryCatch(
                readValue,
                Expression.Catch(
                    typeof(InvalidCastException),
                    Expression.Throw(fallbackException, property.PropertyType)),
                Expression.Catch(
                    typeof(FormatException),
                    Expression.Throw(fallbackException, property.PropertyType)),
                Expression.Catch(
                    typeof(OverflowException),
                    Expression.Throw(fallbackException, property.PropertyType)));
            BinaryExpression assignment = Expression.Assign(
                Expression.Property(target, property),
                value);
            BlockExpression body = Expression.Block(
                new[] { value },
                guardedRead,
                assignment,
                target);
            return Expression.Lambda<Func<T, DbDataReader, int, T>>(body, target, reader, ordinal).Compile();
        } catch {
            return null;
        }
    }

    private static Func<DbDataReader, int, object?>? CreateReaderValueAccessor(PropertyInfo property) {
#if NET8_0_OR_GREATER
        if (!RuntimeFeature.IsDynamicCodeSupported) return null;
#endif
        MethodInfo? getter = GetTypedGetter(property.PropertyType);
        if (getter is null) return null;
        try {
            ParameterExpression reader = Expression.Parameter(typeof(DbDataReader), "reader");
            ParameterExpression ordinal = Expression.Parameter(typeof(int), "ordinal");
            UnaryExpression value = Expression.Convert(
                Expression.Call(reader, getter, ordinal),
                typeof(object));
            return Expression.Lambda<Func<DbDataReader, int, object?>>(value, reader, ordinal).Compile();
        } catch {
            return null;
        }
    }

    private static Func<DbDataReader, T>? CreateFastReaderMap(MappingBinding[] bindings) {
#if NET8_0_OR_GREATER
        if (!RuntimeFeature.IsDynamicCodeSupported) return null;
#endif
        var getters = new MethodInfo[bindings.Length];
        for (int index = 0; index < bindings.Length; index++) {
            MethodInfo? getter = GetTypedGetter(bindings[index].Property.ValueType);
            if (getter is null) return null;
            getters[index] = getter;
        }

        try {
            ParameterExpression reader = Expression.Parameter(typeof(DbDataReader), "reader");
            ParameterExpression instance = Expression.Variable(typeof(T), "instance");
            var values = new ParameterExpression[bindings.Length];
            var body = new List<Expression>((bindings.Length * 2) + 2);

            for (int index = 0; index < bindings.Length; index++) {
                MappingBinding binding = bindings[index];
                ParameterExpression value = Expression.Variable(binding.Property.ValueType, $"value{index}");
                values[index] = value;
                BinaryExpression readValue = Expression.Assign(
                    value,
                    Expression.Call(reader, getters[index], Expression.Constant(binding.ColumnIndex)));
                NewExpression fallbackException = Expression.New(typeof(TypedReaderValueException));
                body.Add(Expression.TryCatch(
                    readValue,
                    Expression.Catch(
                        typeof(InvalidCastException),
                        Expression.Throw(fallbackException, binding.Property.ValueType)),
                    Expression.Catch(
                        typeof(FormatException),
                        Expression.Throw(fallbackException, binding.Property.ValueType)),
                    Expression.Catch(
                        typeof(OverflowException),
                        Expression.Throw(fallbackException, binding.Property.ValueType))));
            }

            body.Add(Expression.Assign(instance, Expression.New(typeof(T))));
            for (int index = 0; index < bindings.Length; index++) {
                body.Add(Expression.Assign(
                    Expression.Property(instance, bindings[index].Property.Property),
                    values[index]));
            }
            body.Add(instance);

            var variables = new ParameterExpression[values.Length + 1];
            variables[0] = instance;
            Array.Copy(values, 0, variables, 1, values.Length);
            return Expression.Lambda<Func<DbDataReader, T>>(
                Expression.Block(variables, body),
                reader).Compile();
        } catch {
            return null;
        }
    }

    internal sealed class TypedReaderValueException : Exception {
    }

    private static MethodInfo? GetTypedGetter(Type type) {
        string? getterName = type == typeof(string) ? nameof(DbDataReader.GetString)
            : type == typeof(bool) ? nameof(DbDataReader.GetBoolean)
            : type == typeof(byte) ? nameof(DbDataReader.GetByte)
            : type == typeof(char) ? nameof(DbDataReader.GetChar)
            : type == typeof(DateTime) ? nameof(DbDataReader.GetDateTime)
            : type == typeof(decimal) ? nameof(DbDataReader.GetDecimal)
            : type == typeof(double) ? nameof(DbDataReader.GetDouble)
            : type == typeof(float) ? nameof(DbDataReader.GetFloat)
            : type == typeof(Guid) ? nameof(DbDataReader.GetGuid)
            : type == typeof(short) ? nameof(DbDataReader.GetInt16)
            : type == typeof(int) ? nameof(DbDataReader.GetInt32)
            : type == typeof(long) ? nameof(DbDataReader.GetInt64)
            : null;
        return getterName is null
            ? null
            : typeof(DbDataReader).GetMethod(getterName, new[] { typeof(int) });
    }

    private sealed class CacheEntry {
        internal CacheEntry(
            string[] headers,
            bool requireAllColumnsMapped,
            AutomaticRowMappingPlan<T> plan) {
            Headers = headers;
            RequireAllColumnsMapped = requireAllColumnsMapped;
            Plan = plan;
        }

        private string[] Headers { get; }
        private bool RequireAllColumnsMapped { get; }
        internal AutomaticRowMappingPlan<T> Plan { get; }

        internal bool Matches(IReadOnlyList<string> headers, bool requireAllColumnsMapped) {
            if (RequireAllColumnsMapped != requireAllColumnsMapped || Headers.Length != headers.Count) {
                return false;
            }

            for (int index = 0; index < Headers.Length; index++) {
                if (!string.Equals(Headers[index], headers[index], StringComparison.Ordinal)) return false;
            }
            return true;
        }
    }
}
