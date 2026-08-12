using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Data {
    public partial class ObjectFlattener {
        private static void FlattenDictionary(
            IReadOnlyList<ObjectDictionaryEntry> entries,
            Dictionary<string, object?> dict,
            string prefix,
            int depth,
            ObjectFlattenerOptions opts,
            HashSet<object> activeObjects) {
            foreach (ObjectDictionaryEntry entry in entries) {
                string? name = entry.Key?.ToString();
                if (string.IsNullOrWhiteSpace(name)) continue;

                string path = string.IsNullOrEmpty(prefix) ? name! : prefix + "." + name;
                if (ShouldIgnorePath(path, opts.Ignore)) continue;
                bool materializePath = IsPathSelectedForMaterialization(path, opts);
                bool expand = opts.ExpandProperties.Contains(name!) || opts.ExpandProperties.Contains(path);
                bool traversePath = expand && CanContainSelectedDescendant(path, opts);
                if (!materializePath && !traversePath) continue;

                object? value = entry.Value;
                if (value == null) {
                    if (materializePath) {
                        SetColumnValue(dict, path, ApplyNullPolicy(path, null, opts), opts);
                    }
                    continue;
                }

                if (expand && !IsSimple(value.GetType()) && depth + 1 < opts.MaxDepth) {
                    if (opts.IncludeFullObjects && materializePath) SetColumnValue(dict, path, value, opts);
                    if (traversePath) FlattenInternal(value, dict, path, depth + 1, opts, activeObjects);
                    continue;
                }

                if (value is IEnumerable enumerable && value is not string) {
                    if (materializePath) {
                        EnsureColumnCapacity(dict, path, opts);
                        SetColumnValue(dict, path, HandleCollection(path, enumerable, opts), opts);
                    }
                    continue;
                }

                if (materializePath) {
                    SetColumnValue(dict, path, ApplyFormatting(path, value, opts), opts);
                }
            }
        }
    }
}
