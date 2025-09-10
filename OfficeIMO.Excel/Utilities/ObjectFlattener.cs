using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Configuration for flattening objects into key/value pairs where keys represent dotted paths.
    /// </summary>
    public class ObjectFlattenerOptions {
        /// <summary>Property names or dotted paths to expand (descend into) rather than treat as simple values.</summary>
        public List<string> ExpandProperties { get; } = new();
        /// <summary>When true, includes the original object under its path in addition to expanded fields.</summary>
        public bool IncludeFullObjects { get; set; }
        /// <summary>Maximum recursion depth when expanding nested objects.</summary>
        public int MaxDepth { get; set; } = int.MaxValue;
        /// <summary>Header casing strategy for generated column names.</summary>
        public HeaderCase HeaderCase { get; set; } = HeaderCase.Raw;
        /// <summary>Optional prefixes to trim from generated headers.</summary>
        public string[] HeaderPrefixTrimPaths { get; set; } = Array.Empty<string>();
        /// <summary>Optional explicit column whitelist (dotted paths).</summary>
        public string[]? Columns { get; set; }
        /// <summary>Paths to exclude from output.</summary>
        public string[] Ignore { get; set; } = Array.Empty<string>();
        /// <summary>
        /// Optional list of dotted paths to pin to the front of the generated column order when <see cref="Columns"/> is not specified.
        /// Matching paths keep the specified order; any non-matching paths follow in natural order.
        /// </summary>
        public string[] PinnedFirst { get; set; } = Array.Empty<string>();
        /// <summary>How null values are represented.</summary>
        public NullPolicy NullPolicy { get; set; } = NullPolicy.NullLiteral;
        /// <summary>Per‑path default values used when <see cref="NullPolicy.DefaultValue"/> is selected.</summary>
        public Dictionary<string, object?> DefaultValues { get; } = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>Per‑path formatting delegates applied to values.</summary>
        public Dictionary<string, Func<object?, object?>> Formatters { get; } = new(StringComparer.OrdinalIgnoreCase);
        /// <summary>How to handle collections.</summary>
        public CollectionMode CollectionMode { get; set; } = CollectionMode.JoinWith;
        /// <summary>Delimiter used when <see cref="CollectionMode.JoinWith"/> is selected.</summary>
        public string CollectionJoinWith { get; set; } = ",";

        /// <summary>
        /// Maps collection paths (e.g., "ScoreBreakdown") to dynamic columns using element properties.
        /// Example: map path "ScoreBreakdown" with KeyProperty="Name", ValueProperty="Value" to produce columns like
        /// "ScoreBreakdown.HasMX" = 2, "ScoreBreakdown.EffectiveSPFSends" = 2.
        /// </summary>
        public System.Collections.Generic.Dictionary<string, CollectionColumnMapping> CollectionMapColumns { get; } = new(System.StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Flattens objects to a dictionary of dotted-path keys to values suitable for table generation.
    /// </summary>
    public class ObjectFlattener {
        private static readonly ConcurrentDictionary<Type, PropertyInfo[]> _cache = new();

        /// <summary>
        /// Flattens <paramref name="item"/> into a dictionary according to <paramref name="opts"/>.
        /// </summary>
        public Dictionary<string, object?> Flatten<T>(T item, ObjectFlattenerOptions opts) {
            var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            if (item == null) return result;
            FlattenInternal(item!, result, string.Empty, 0, opts);
            return result;
        }

        /// <summary>
        /// Computes all reachable dotted paths for a given <paramref name="type"/> under <paramref name="opts"/>.
        /// </summary>
        public List<string> GetPaths(Type type, ObjectFlattenerOptions opts) {
            var paths = new List<string>();
            BuildPaths(type, string.Empty, 0, opts, paths);
            return paths
                .Where(p => !opts.Ignore.Any(i => p.StartsWith(i, StringComparison.OrdinalIgnoreCase)))
                .ToList();
        }

        private static void FlattenInternal(object obj, Dictionary<string, object?> dict, string prefix, int depth, ObjectFlattenerOptions opts) {
            if (depth >= opts.MaxDepth) return;

            var type = obj.GetType();

            // Special-case: ValueTuple (struct tuples) expose public fields (Item1..ItemN) not properties.
            if (IsValueTuple(type))
            {
                FlattenValueTuple(obj, dict, prefix, depth, opts);
                return;
            }

            var props = _cache.GetOrAdd(type, t => t.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                                                .OrderBy(p => p.MetadataToken).ToArray());
            foreach (var prop in props) {
                var value = prop.GetValue(obj);
                var path = string.IsNullOrEmpty(prefix) ? prop.Name : prefix + "." + prop.Name;
                if (opts.Ignore.Any(i => path.StartsWith(i, StringComparison.OrdinalIgnoreCase))) continue;

                bool expand = opts.ExpandProperties.Contains(prop.Name) || opts.ExpandProperties.Contains(path);
                bool isCollection = value is IEnumerable && value is not string;

                if (value == null) {
                    dict[path] = ApplyNullPolicy(path, null, opts);
                    continue;
                }

                if (isCollection) {
                    if (opts.CollectionMapColumns.TryGetValue(path, out var map)) {
                        MapCollectionToColumns(path, (IEnumerable)value, map, dict, opts);
                    } else {
                        dict[path] = HandleCollection(path, (IEnumerable)value, opts);
                    }
                    continue;
                }

                if (!expand || IsSimple(prop.PropertyType)) {
                    dict[path] = ApplyFormatting(path, value, opts);
                    continue;
                }

                if (opts.IncludeFullObjects) {
                    dict[path] = value;
                }

                if (depth + 1 < opts.MaxDepth) {
                    FlattenInternal(value, dict, path, depth + 1, opts);
                }
            }
        }

        private static bool IsValueTuple(Type t)
        {
            if (!t.IsValueType) return false;
            var n = t.Name;
            if (n.StartsWith("ValueTuple`", StringComparison.Ordinal)) return true;
            // non-generic System.ValueTuple
            return string.Equals(t.FullName, "System.ValueTuple", StringComparison.Ordinal);
        }

        private static void FlattenValueTuple(object obj, Dictionary<string, object?> dict, string prefix, int depth, ObjectFlattenerOptions opts)
        {
            // Try ITuple via reflection (available on newer frameworks). Avoids compile-time dependency for netstandard2.0
            var iTupleType = Type.GetType("System.Runtime.CompilerServices.ITuple");
            if (iTupleType != null && iTupleType.IsAssignableFrom(obj.GetType()))
            {
                var lenProp = iTupleType.GetProperty("Length");
                var itemProp = iTupleType.GetProperty("Item"); // indexer
                if (lenProp != null && itemProp != null)
                {
                    int length = Convert.ToInt32(lenProp.GetValue(obj, null));
                    for (int i = 0; i < length; i++)
                    {
                        var path = string.IsNullOrEmpty(prefix) ? $"Item{i + 1}" : $"{prefix}.Item{i + 1}";
                        var val = itemProp.GetValue(obj, new object[] { i });
                        if (val == null)
                        {
                            dict[path] = ApplyNullPolicy(path, null, opts);
                        }
                        else if (IsSimple(val.GetType()))
                        {
                            dict[path] = ApplyFormatting(path, val, opts);
                        }
                        else
                        {
                            FlattenInternal(val, dict, path, depth + 1, opts);
                        }
                    }
                    return;
                }
            }

            // Fallback: reflect public instance fields Item1..ItemN
            var fields = obj.GetType().GetFields(BindingFlags.Public | BindingFlags.Instance)
                .Where(f => f.Name.StartsWith("Item", StringComparison.Ordinal))
                .OrderBy(f => f.Name, StringComparer.Ordinal)
                .ToArray();
            int idx = 1;
            foreach (var f in fields)
            {
                var path = string.IsNullOrEmpty(prefix) ? $"Item{idx}" : $"{prefix}.Item{idx}";
                var val = f.GetValue(obj);
                if (val == null)
                {
                    dict[path] = ApplyNullPolicy(path, null, opts);
                }
                else if (IsSimple(val.GetType()))
                {
                    dict[path] = ApplyFormatting(path, val, opts);
                }
                else
                {
                    FlattenInternal(val, dict, path, depth + 1, opts);
                }
                idx++;
            }
        }

        private static void MapCollectionToColumns(string basePath, IEnumerable enumerable, CollectionColumnMapping map, Dictionary<string, object?> dict, ObjectFlattenerOptions opts)
        {
            foreach (var item in enumerable)
            {
                if (item == null) continue;
                var t = item.GetType();
                var keyProp = t.GetProperty(map.KeyProperty);
                var valProp = t.GetProperty(map.ValueProperty);
                if (keyProp == null || valProp == null) continue;

                var keyObj = keyProp.GetValue(item);
                if (keyObj == null) continue;
                var key = keyObj.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(key)) continue;

                var colPath = basePath + "." + key;
                if (opts.Ignore.Any(i => colPath.StartsWith(i, StringComparison.OrdinalIgnoreCase))) continue;
                var value = valProp.GetValue(item);
                dict[colPath] = ApplyFormatting(colPath, value, opts);
            }
        }

        private static void BuildPaths(Type type, string prefix, int depth, ObjectFlattenerOptions opts, List<string> paths) {
            if (depth >= opts.MaxDepth) return;
            if (IsValueTuple(type))
            {
                // Prefer counting actual Item* fields for precision (covers non-generic System.ValueTuple)
                int itemCount = objFieldCount(type);
                // If field count is 0 but the type is generic, fall back to generic arity (covers ITuple-backed cases)
                if (itemCount == 0 && type.IsGenericType)
                    itemCount = type.GetGenericArguments().Length;

                for (int i = 1; i <= itemCount; i++)
                {
                    var path = string.IsNullOrEmpty(prefix) ? $"Item{i}" : $"{prefix}.Item{i}";
                    if (!opts.Ignore.Any(x => path.StartsWith(x, StringComparison.OrdinalIgnoreCase)))
                        paths.Add(path);
                }
                return;
            }
            var props = _cache.GetOrAdd(type, t => t.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                                   .OrderBy(p => p.MetadataToken).ToArray());
            foreach (var prop in props) {
                var path = string.IsNullOrEmpty(prefix) ? prop.Name : prefix + "." + prop.Name;
                if (opts.Ignore.Any(i => path.StartsWith(i, StringComparison.OrdinalIgnoreCase))) continue;
                bool expand = opts.ExpandProperties.Contains(prop.Name) || opts.ExpandProperties.Contains(path);
                bool isCollection = typeof(IEnumerable).IsAssignableFrom(prop.PropertyType) && prop.PropertyType != typeof(string);
                if (isCollection) {
                    paths.Add(path);
                    continue;
                }
                if (!expand || opts.IncludeFullObjects || IsSimple(prop.PropertyType)) {
                    paths.Add(path);
                }
                if (expand && !IsSimple(prop.PropertyType)) {
                    BuildPaths(prop.PropertyType, path, depth + 1, opts, paths);
                }
            }
        }

        private static int objFieldCount(Type valueTupleType)
        {
            try
            {
                return valueTupleType
                    .GetFields(BindingFlags.Public | BindingFlags.Instance)
                    .Count(f => f.Name.StartsWith("Item", StringComparison.Ordinal));
            }
            catch { return 0; }
        }

        private static object? HandleCollection(string path, IEnumerable enumerable, ObjectFlattenerOptions opts) {
            if (opts.CollectionMode == CollectionMode.JoinWith) {
                var list = enumerable.Cast<object?>().Select(v => v?.ToString()).ToArray();
                var joined = string.Join(opts.CollectionJoinWith, list);
                return ApplyFormatting(path, joined, opts);
            }
            // ExpandRows handled in SheetBuilder
            return enumerable;
        }

        private static object? ApplyFormatting(string path, object? value, ObjectFlattenerOptions opts) {
            if (value == null) return ApplyNullPolicy(path, null, opts);
            if (opts.Formatters.TryGetValue(path, out var formatter)) {
                return formatter(value);
            }
            return value;
        }

        private static object? ApplyNullPolicy(string path, object? value, ObjectFlattenerOptions opts) {
            return opts.NullPolicy switch {
                NullPolicy.EmptyString => string.Empty,
                NullPolicy.DefaultValue => opts.DefaultValues.TryGetValue(path, out var v) ? v : null,
                _ => null
            };
        }

        private static bool IsSimple(Type type) {
            return type.IsPrimitive || type.IsEnum || type == typeof(string) || type == typeof(decimal) || type == typeof(DateTime) || type == typeof(DateTimeOffset) || type == typeof(TimeSpan) || type == typeof(Guid);
        }
    }
}
    /// <summary>
    /// Configuration for mapping a collection of objects into dynamic columns.
    /// </summary>
public sealed class CollectionColumnMapping
{
        /// <summary>Property name on the item to use as the header suffix (key).</summary>
        public string KeyProperty { get; set; } = "Name";
        /// <summary>Property name on the item to use as the cell value.</summary>
        public string ValueProperty { get; set; } = "Value";
        /// <summary>Optional prefix for generated column headers.</summary>
        public string? HeaderPrefix { get; set; }
}
