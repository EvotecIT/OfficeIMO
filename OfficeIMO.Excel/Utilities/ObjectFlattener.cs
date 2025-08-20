using System;
using System.Collections.Generic;
using System.Reflection;

namespace OfficeIMO.Excel.Utilities {
    public class ObjectFlattenerOptions {
        public List<string> ExpandProperties { get; } = new List<string>();
        public bool IncludeFullObjects { get; set; }
        public int MaxDepth { get; set; } = int.MaxValue;
    }

    public class ObjectFlattener {
        public Dictionary<string, object?> Flatten<T>(T item, ObjectFlattenerOptions opts) {
            var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            if (item == null) return result;
            FlattenInternal(item!, result, string.Empty, 0, opts);
            return result;
        }

        private static void FlattenInternal(object obj, Dictionary<string, object?> dict, string prefix, int depth, ObjectFlattenerOptions opts) {
            if (depth >= opts.MaxDepth) return;
            var props = obj.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (var prop in props) {
                var value = prop.GetValue(obj);
                var path = string.IsNullOrEmpty(prefix) ? prop.Name : prefix + "." + prop.Name;
                bool expand = opts.ExpandProperties.Contains(prop.Name) || opts.ExpandProperties.Contains(path);
                if (value == null || IsSimple(prop.PropertyType) || !expand) {
                    dict[path] = value;
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

        private static bool IsSimple(Type type) {
            return type.IsPrimitive || type.IsEnum || type == typeof(string) || type == typeof(decimal) || type == typeof(DateTime) || type == typeof(DateTimeOffset) || type == typeof(TimeSpan) || type == typeof(Guid);
        }
    }
}