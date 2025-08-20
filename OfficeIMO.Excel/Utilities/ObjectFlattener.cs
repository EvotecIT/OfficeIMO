using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.Json;

namespace OfficeIMO.Excel.Utilities {
    public enum HeaderCase {
        Raw,
        Pascal,
        Title
    }

    public enum NullPolicy {
        EmptyString,
        NullLiteral,
        DefaultValue
    }

    public enum CollectionMode {
        JoinWith,
        ExpandRows,
        Json
    }

    public class ObjectFlattenerOptions {
        public List<string> ExpandProperties { get; } = new();
        public bool IncludeFullObjects { get; set; }
        public int MaxDepth { get; set; } = int.MaxValue;
        public HeaderCase HeaderCase { get; set; } = HeaderCase.Raw;
        public string[] HeaderPrefixTrimPaths { get; set; } = Array.Empty<string>();
        public string[]? Columns { get; set; }
        public string[] Ignore { get; set; } = Array.Empty<string>();
        public NullPolicy NullPolicy { get; set; } = NullPolicy.NullLiteral;
        public Dictionary<string, object?> DefaultValues { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, Func<object?, object?>> Formatters { get; } = new(StringComparer.OrdinalIgnoreCase);
        public CollectionMode CollectionMode { get; set; } = CollectionMode.JoinWith;
        public string CollectionJoinWith { get; set; } = ",";
    }

    public class ObjectFlattener {
        private static readonly ConcurrentDictionary<Type, PropertyInfo[]> _cache = new();

        public Dictionary<string, object?> Flatten<T>(T item, ObjectFlattenerOptions opts) {
            var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            if (item == null) return result;
            FlattenInternal(item!, result, string.Empty, 0, opts);
            return result;
        }

        public List<string> GetPaths(Type type, ObjectFlattenerOptions opts) {
            var paths = new List<string>();
            BuildPaths(type, string.Empty, 0, opts, paths);
            return paths
                .Where(p => !opts.Ignore.Any(i => p.StartsWith(i, StringComparison.OrdinalIgnoreCase)))
                .ToList();
        }

        private static void FlattenInternal(object obj, Dictionary<string, object?> dict, string prefix, int depth, ObjectFlattenerOptions opts) {
            if (depth >= opts.MaxDepth) return;

            var props = _cache.GetOrAdd(obj.GetType(), t => t.GetProperties(BindingFlags.Public | BindingFlags.Instance)
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
                    dict[path] = HandleCollection(path, (IEnumerable)value, opts);
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

        private static void BuildPaths(Type type, string prefix, int depth, ObjectFlattenerOptions opts, List<string> paths) {
            if (depth >= opts.MaxDepth) return;
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

        private static object? HandleCollection(string path, IEnumerable enumerable, ObjectFlattenerOptions opts) {
            if (opts.CollectionMode == CollectionMode.JoinWith) {
                var list = enumerable.Cast<object?>().Select(v => v?.ToString()).ToArray();
                var joined = string.Join(opts.CollectionJoinWith, list);
                return ApplyFormatting(path, joined, opts);
            }
            if (opts.CollectionMode == CollectionMode.Json) {
                return JsonSerializer.Serialize(enumerable);
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
