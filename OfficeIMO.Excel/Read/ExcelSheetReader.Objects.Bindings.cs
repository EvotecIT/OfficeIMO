using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Threading;
using System.Text;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Runtime.Serialization;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Object-mapping readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private sealed class TypedPropertyBinding<TTarget> {
            internal TypedPropertyBinding(
                PropertyInfo property,
                Type propertyType,
                Type destinationType,
                TypedBindingKind bindingKind,
                bool isNullable,
                bool needsDateStyleConversion,
                Action<TTarget, object?> setValue,
                Action<TTarget, string?>? setString,
                Action<TTarget, int>? setInt32,
                Action<TTarget, long>? setInt64,
                Action<TTarget, double>? setDouble,
                Action<TTarget, decimal>? setDecimal,
                Action<TTarget, bool>? setBoolean,
                Action<TTarget, DateTime>? setDateTime,
                Func<object, CultureInfo, object?> convertValue) {
                Property = property;
                PropertyType = propertyType;
                DestinationType = destinationType;
                BindingKind = bindingKind;
                IsNullable = isNullable;
                NeedsDateStyleConversion = needsDateStyleConversion;
                SetValue = setValue;
                SetString = setString;
                SetInt32 = setInt32;
                SetInt64 = setInt64;
                SetDouble = setDouble;
                SetDecimal = setDecimal;
                SetBoolean = setBoolean;
                SetDateTime = setDateTime;
                ConvertValue = convertValue;
            }

            internal PropertyInfo Property { get; }
            internal Type PropertyType { get; }
            internal Type DestinationType { get; }
            internal TypedBindingKind BindingKind { get; }
            internal bool IsNullable { get; }
            internal bool NeedsDateStyleConversion { get; }
            internal Action<TTarget, object?> SetValue { get; }
            internal Action<TTarget, string?>? SetString { get; }
            internal Action<TTarget, int>? SetInt32 { get; }
            internal Action<TTarget, long>? SetInt64 { get; }
            internal Action<TTarget, double>? SetDouble { get; }
            internal Action<TTarget, decimal>? SetDecimal { get; }
            internal Action<TTarget, bool>? SetBoolean { get; }
            internal Action<TTarget, DateTime>? SetDateTime { get; }
            internal Func<object, CultureInfo, object?> ConvertValue { get; }
        }

        private enum TypedBindingKind {
            Other,
            String,
            Int32,
            Int64,
            Double,
            Decimal,
            Boolean,
            DateTime
        }

        private sealed class TypedPropertyMapCache {
            internal TypedPropertyMapCache(
                string typeName,
                Dictionary<string, PropertyInfo> exactProperties,
                Dictionary<string, PropertyInfo> exactAliases,
                Dictionary<string, PropertyInfo> canonicalProperties,
                Dictionary<string, PropertyInfo> canonicalAliases,
                IReadOnlyList<string> diagnostics) {
                TypeName = typeName;
                ExactProperties = exactProperties;
                ExactAliases = exactAliases;
                CanonicalProperties = canonicalProperties;
                CanonicalAliases = canonicalAliases;
                Diagnostics = diagnostics;
            }

            internal string TypeName { get; }
            internal Dictionary<string, PropertyInfo> ExactProperties { get; }
            internal Dictionary<string, PropertyInfo> ExactAliases { get; }
            internal Dictionary<string, PropertyInfo> CanonicalProperties { get; }
            internal Dictionary<string, PropertyInfo> CanonicalAliases { get; }
            internal IReadOnlyList<string> Diagnostics { get; }
        }

        private sealed class TypedHeaderBindingCache<TTarget> where TTarget : new() {
            internal TypedHeaderBindingCache(TypedPropertyBinding<TTarget>?[] bindings, IReadOnlyList<string> unmappedIssues) {
                Bindings = bindings;
                UnmappedIssues = unmappedIssues;
            }

            internal TypedPropertyBinding<TTarget>?[] Bindings { get; }
            internal IReadOnlyList<string> UnmappedIssues { get; }
        }

        private static class TypedObjectBindingCache<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] TTarget> where TTarget : new() {
            private const int HeaderBindingCacheLimit = 64;
            private const int HeaderBindingCacheCharacterLimit = 65_536;

            internal static readonly PropertyInfo[] WritableProperties = typeof(TTarget)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => p.CanWrite)
                .ToArray();

            internal static readonly Dictionary<PropertyInfo, TypedPropertyBinding<TTarget>> Bindings = CreateBindings();

            internal static readonly TypedPropertyMapCache PropertyMaps = CreatePropertyMaps();

            private static readonly TypedHeaderBindingCache<TTarget> WritablePropertyOrderBindings = CreateWritablePropertyOrderBindings();

            private static readonly ConcurrentDictionary<string, TypedHeaderBindingCache<TTarget>> HeaderBindings = new ConcurrentDictionary<string, TypedHeaderBindingCache<TTarget>>(StringComparer.Ordinal);

            internal static TypedHeaderBindingCache<TTarget> GetHeaderBindings(string[] headers) {
                if (HeadersMatchWritablePropertyOrder(headers)) {
                    return WritablePropertyOrderBindings;
                }

                if (!CanCacheHeaders(headers)) {
                    return CreateHeaderBindings(headers);
                }

                string key = CreateHeaderBindingKey(headers);
                if (HeaderBindings.TryGetValue(key, out var cached)) {
                    return cached;
                }

                var created = CreateHeaderBindings(headers);
                if (HeaderBindings.Count < HeaderBindingCacheLimit) {
                    return HeaderBindings.GetOrAdd(key, created);
                }

                return created;
            }

            private static bool CanCacheHeaders(string[] headers) {
                long characters = 0;
                for (int index = 0; index < headers.Length; index++) {
                    characters += headers[index]?.Length ?? 0;
                    if (characters > HeaderBindingCacheCharacterLimit) {
                        return false;
                    }
                }

                return true;
            }

            private static Dictionary<PropertyInfo, TypedPropertyBinding<TTarget>> CreateBindings() {
                var bindings = new Dictionary<PropertyInfo, TypedPropertyBinding<TTarget>>(WritableProperties.Length);
                foreach (var property in WritableProperties) {
                    bindings.Add(property, CreateBinding(property));
                }

                return bindings;
            }

            private static TypedHeaderBindingCache<TTarget> CreateWritablePropertyOrderBindings() {
                var map = new TypedPropertyBinding<TTarget>?[WritableProperties.Length];
                for (int i = 0; i < WritableProperties.Length; i++) {
                    map[i] = Bindings[WritableProperties[i]];
                }

                return new TypedHeaderBindingCache<TTarget>(map, Array.Empty<string>());
            }

            private static bool HeadersMatchWritablePropertyOrder(string[] headers) {
                if (headers.Length != WritableProperties.Length) {
                    return false;
                }

                for (int i = 0; i < headers.Length; i++) {
                    if (!string.Equals(headers[i], WritableProperties[i].Name, StringComparison.Ordinal)) {
                        return false;
                    }
                }

                return true;
            }

            private static TypedPropertyBinding<TTarget> CreateBinding(PropertyInfo property) {
                var nullable = Nullable.GetUnderlyingType(property.PropertyType);
                var destinationType = nullable ?? property.PropertyType;
                var bindingKind = GetBindingKind(destinationType);
                return new TypedPropertyBinding<TTarget>(
                    property,
                    property.PropertyType,
                    destinationType,
                    bindingKind,
                    !property.PropertyType.IsValueType || nullable != null,
                    NeedsDateStyleConversion(destinationType),
                    CreateSetter(property),
                    bindingKind == TypedBindingKind.String ? CreateTypedSetter<string?>(property) : null,
                    bindingKind == TypedBindingKind.Int32 ? CreateTypedSetter<int>(property) : null,
                    bindingKind == TypedBindingKind.Int64 ? CreateTypedSetter<long>(property) : null,
                    bindingKind == TypedBindingKind.Double ? CreateTypedSetter<double>(property) : null,
                    bindingKind == TypedBindingKind.Decimal ? CreateTypedSetter<decimal>(property) : null,
                    bindingKind == TypedBindingKind.Boolean ? CreateTypedSetter<bool>(property) : null,
                    bindingKind == TypedBindingKind.DateTime ? CreateTypedSetter<DateTime>(property) : null,
                    CreateConverter(destinationType));
            }

            private static TypedBindingKind GetBindingKind(Type destinationType) {
                if (destinationType == typeof(string)) {
                    return TypedBindingKind.String;
                }

                if (destinationType == typeof(int)) {
                    return TypedBindingKind.Int32;
                }

                if (destinationType == typeof(long)) {
                    return TypedBindingKind.Int64;
                }

                if (destinationType == typeof(double)) {
                    return TypedBindingKind.Double;
                }

                if (destinationType == typeof(decimal)) {
                    return TypedBindingKind.Decimal;
                }

                if (destinationType == typeof(bool)) {
                    return TypedBindingKind.Boolean;
                }

                if (destinationType == typeof(DateTime)) {
                    return TypedBindingKind.DateTime;
                }

                return TypedBindingKind.Other;
            }

            private static bool NeedsDateStyleConversion(Type destinationType) {
                return destinationType == typeof(DateTime) || destinationType == typeof(string);
            }

            private static Func<object, CultureInfo, object?> CreateConverter(Type destinationType) {
                if (destinationType == typeof(string)) {
                    return static (value, culture) => value as string ?? Convert.ToString(value, culture);
                }

                if (destinationType == typeof(bool)) {
                    return static (value, culture) => {
                        try {
                            if (value is bool boolValue) return boolValue;
                            return Convert.ToBoolean(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(int)) {
                    return static (value, culture) => {
                        try {
                            if (value is int intValue) return intValue;
                            if (value is double doubleValue
                                && doubleValue >= int.MinValue
                                && doubleValue <= int.MaxValue
                                && Math.Truncate(doubleValue) == doubleValue) {
                                return (int)doubleValue;
                            }

                            return Convert.ToInt32(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(long)) {
                    return static (value, culture) => {
                        try {
                            if (value is long longValue) return longValue;
                            if (value is double doubleValue
                                && doubleValue >= long.MinValue
                                && doubleValue <= long.MaxValue
                                && Math.Truncate(doubleValue) == doubleValue) {
                                return (long)doubleValue;
                            }

                            return Convert.ToInt64(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(double)) {
                    return static (value, culture) => {
                        try {
                            if (value is double doubleValue) return doubleValue;
                            return Convert.ToDouble(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(decimal)) {
                    return static (value, culture) => {
                        try {
                            if (value is decimal decimalValue) return decimalValue;
                            return Convert.ToDecimal(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(DateTime)) {
                    return static (value, culture) => {
                        try {
                            if (value is DateTime dt) return dt;
                            if (value is double oa) return DateTime.FromOADate(oa);
                            if (DateTime.TryParse(Convert.ToString(value, culture), culture, DateTimeStyles.AssumeLocal, out var parsed)) return parsed;
                            return null;
                        } catch {
                            return null;
                        }
                    };
                }

                return (value, culture) => ConvertToDestinationType(value, destinationType, culture);
            }

            private static Action<TTarget, object?> CreateSetter(PropertyInfo property) {
                try {
                    var target = Expression.Parameter(typeof(TTarget), "target");
                    var value = Expression.Parameter(typeof(object), "value");
                    var converted = Expression.Convert(value, property.PropertyType);
                    var body = Expression.Assign(Expression.Property(target, property), converted);
                    return Expression.Lambda<Action<TTarget, object?>>(body, target, value).Compile();
                } catch {
                    return (target, value) => property.SetValue(target, value);
                }
            }

            private static Action<TTarget, TValue>? CreateTypedSetter<TValue>(PropertyInfo property) {
                try {
                    var target = Expression.Parameter(typeof(TTarget), "target");
                    var value = Expression.Parameter(typeof(TValue), "value");
                    var converted = Expression.Convert(value, property.PropertyType);
                    var body = Expression.Assign(Expression.Property(target, property), converted);
                    return Expression.Lambda<Action<TTarget, TValue>>(body, target, value).Compile();
                } catch {
                    return null;
                }
            }

            private static TypedPropertyMapCache CreatePropertyMaps() {
                string typeName = typeof(TTarget).Name;
                var diagnostics = new List<string>();

                return new TypedPropertyMapCache(
                    typeName,
                    BuildPropertyMap(WritableProperties, prop => new[] { prop.Name }, diagnostics, typeName, "exact property"),
                    BuildPropertyMap(WritableProperties, GetPropertyAliases, diagnostics, typeName, "explicit alias"),
                    BuildPropertyMap(WritableProperties, prop => new[] { CanonicalizeMemberName(prop.Name) }, diagnostics, typeName, "friendly property"),
                    BuildPropertyMap(WritableProperties, prop => GetPropertyAliases(prop).Select(CanonicalizeMemberName), diagnostics, typeName, "friendly alias"),
                    diagnostics.ToArray());
            }

            private static TypedHeaderBindingCache<TTarget> CreateHeaderBindings(string[] headers) {
                var map = new TypedPropertyBinding<TTarget>?[headers.Length];
                int mappedCount = 0;

                // Exact property matches win first so alias/friendly fallback does not steal
                // a property from a later exact-name column.
                for (int c = 0; c < headers.Length; c++) {
                    if (PropertyMaps.ExactProperties.TryGetValue(headers[c], out var pi)) {
                        map[c] = Bindings[pi];
                        mappedCount++;
                    }
                }

                if (mappedCount == headers.Length) {
                    return new TypedHeaderBindingCache<TTarget>(map, Array.Empty<string>());
                }

                var assignedProps = new HashSet<PropertyInfo>();
                for (int c = 0; c < map.Length; c++) {
                    if (map[c] != null && PropertyMaps.ExactProperties.TryGetValue(headers[c], out var pi)) {
                        assignedProps.Add(pi);
                    }
                }

                // Explicit aliases come next (DisplayName/DataMember/ExcelColumn).
                for (int c = 0; c < headers.Length; c++) {
                    if (map[c] != null) {
                        continue;
                    }

                    if (PropertyMaps.ExactAliases.TryGetValue(headers[c], out var pi) && !assignedProps.Contains(pi)) {
                        map[c] = Bindings[pi];
                        assignedProps.Add(pi);
                    }
                }

                for (int c = 0; c < headers.Length; c++) {
                    if (map[c] != null) {
                        continue;
                    }

                    string canonicalHeader = CanonicalizeMemberName(headers[c]);
                    if (canonicalHeader.Length == 0) {
                        continue;
                    }

                    if (PropertyMaps.CanonicalProperties.TryGetValue(canonicalHeader, out var pi) && !assignedProps.Contains(pi)) {
                        map[c] = Bindings[pi];
                        assignedProps.Add(pi);
                        continue;
                    }

                    if (PropertyMaps.CanonicalAliases.TryGetValue(canonicalHeader, out pi) && !assignedProps.Contains(pi)) {
                        map[c] = Bindings[pi];
                        assignedProps.Add(pi);
                    }
                }

                var unmappedIssues = new List<string>();
                for (int c = 0; c < headers.Length; c++) {
                    if (map[c] != null) {
                        continue;
                    }

                    string header = headers[c];
                    if (header.Length == 0) {
                        continue;
                    }

                    unmappedIssues.Add($"[TypedRead UnmappedHeader] Type='{PropertyMaps.TypeName}', header='{header}', column={c + 1}.");
                }

                return new TypedHeaderBindingCache<TTarget>(map, unmappedIssues.ToArray());
            }

            private static string CreateHeaderBindingKey(string[] headers) {
                var builder = new StringBuilder(headers.Length * 16);
                for (int i = 0; i < headers.Length; i++) {
                    string header = headers[i] ?? string.Empty;
                    builder.Append(header.Length.ToString(CultureInfo.InvariantCulture));
                    builder.Append(':');
                    builder.Append(header);
                    builder.Append('|');
                }

                return builder.ToString();
            }
        }
    }
}
