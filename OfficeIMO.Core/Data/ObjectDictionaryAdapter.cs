using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Data {
    /// <summary>
    /// Adapts non-generic and generic dictionary implementations without making
    /// <see cref="ObjectFlattener"/> depend on any one dictionary concrete type.
    /// </summary>
    internal static class ObjectDictionaryAdapter {
        private static readonly ConcurrentDictionary<Type, bool> DictionaryTypeCache = new();
        private static readonly ConcurrentDictionary<Type, DictionaryEntryAccessor> EntryAccessorCache = new();

        internal static bool TryGetEntries(object? value, int maximumItems, out List<ObjectDictionaryEntry> entries) {
            return TryGetEntries(value, maximumItems, includeKey: null, maximumItems,
                deduplicateResolvedKeys: false, out entries);
        }

        internal static bool TryGetEntries(
            object? value,
            int maximumItems,
            Func<object?, bool>? includeKey,
            int maximumEntries,
            bool deduplicateResolvedKeys,
            out List<ObjectDictionaryEntry> entries) {
            entries = new List<ObjectDictionaryEntry>();
            if (value == null) return false;
            if (maximumItems <= 0) throw new ArgumentOutOfRangeException(nameof(maximumItems));
            if (maximumEntries <= 0) throw new ArgumentOutOfRangeException(nameof(maximumEntries));
            int itemCount = 0;
            Dictionary<string, int>? retainedIndexes = deduplicateResolvedKeys
                ? new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
                : null;

            if (value is IDictionary dictionary) {
                foreach (DictionaryEntry entry in dictionary) {
                    AddBounded(entries, new ObjectDictionaryEntry(entry.Key, entry.Value), maximumItems,
                        includeKey, maximumEntries, retainedIndexes, ref itemCount);
                }
                return true;
            }

            if (value is IDictionary<string, object?> genericObjectDictionary) {
                foreach (KeyValuePair<string, object?> entry in genericObjectDictionary) {
                    AddBounded(entries, new ObjectDictionaryEntry(entry.Key, entry.Value), maximumItems,
                        includeKey, maximumEntries, retainedIndexes, ref itemCount);
                }
                return true;
            }

            if (value is IReadOnlyDictionary<string, object?> readOnlyObjectDictionary) {
                foreach (KeyValuePair<string, object?> entry in readOnlyObjectDictionary) {
                    AddBounded(entries, new ObjectDictionaryEntry(entry.Key, entry.Value), maximumItems,
                        includeKey, maximumEntries, retainedIndexes, ref itemCount);
                }
                return true;
            }

            Type type = value.GetType();
            if (!DictionaryTypeCache.GetOrAdd(type, IsGenericDictionaryType) || value is not IEnumerable enumerable) {
                return false;
            }

            foreach (object? item in enumerable) {
                if (item == null) continue;
                DictionaryEntryAccessor accessor = EntryAccessorCache.GetOrAdd(item.GetType(), CreateEntryAccessor);
                if (!accessor.IsValid) {
                    entries.Clear();
                    return false;
                }

                AddBounded(entries,
                    new ObjectDictionaryEntry(accessor.Key!.GetValue(item), accessor.Value!.GetValue(item)),
                    maximumItems, includeKey, maximumEntries, retainedIndexes, ref itemCount);
            }

            return true;
        }

        internal static bool IsDictionaryType(Type type) {
            if (type == null) throw new ArgumentNullException(nameof(type));
            return typeof(IDictionary).IsAssignableFrom(type)
                || DictionaryTypeCache.GetOrAdd(type, IsGenericDictionaryType);
        }

        private static void AddBounded(
            List<ObjectDictionaryEntry> entries,
            ObjectDictionaryEntry entry,
            int maximumItems,
            Func<object?, bool>? includeKey,
            int maximumEntries,
            Dictionary<string, int>? retainedIndexes,
            ref int itemCount) {
            if (itemCount++ >= maximumItems) {
                throw new InvalidDataException($"The dictionary exceeds the {maximumItems}-item flattening limit.");
            }
            if (includeKey != null && !includeKey(entry.Key)) return;
            string? resolvedKey = null;
            if (retainedIndexes != null) {
                resolvedKey = entry.Key?.ToString();
                if (string.IsNullOrWhiteSpace(resolvedKey)) return;
                if (retainedIndexes.TryGetValue(resolvedKey!, out int retainedIndex)) {
                    ObjectDictionaryEntry retained = entries[retainedIndex];
                    entries[retainedIndex] = new ObjectDictionaryEntry(retained.Key, entry.Value);
                    return;
                }
            }
            if (entries.Count >= maximumEntries) {
                throw ObjectFlattener.CreateRawColumnLimitException(
                    "Dictionary flattening",
                    checked(entries.Count + 1),
                    maximumEntries);
            }
            if (retainedIndexes != null) retainedIndexes.Add(resolvedKey!, entries.Count);
            entries.Add(entry);
        }

        [UnconditionalSuppressMessage(
            "Trimming",
            "IL2070",
            Justification = "Implemented generic dictionary interfaces are required for runtime interface dispatch and were verified by the published NativeAOT dictionary scenarios.")]
        private static bool IsGenericDictionaryType(Type type) {
            return type.GetInterfaces().Any(interfaceType => {
                if (!interfaceType.IsGenericType) return false;
                Type definition = interfaceType.GetGenericTypeDefinition();
                return definition == typeof(IDictionary<,>) || definition == typeof(IReadOnlyDictionary<,>);
            });
        }

        [DynamicDependency(DynamicallyAccessedMemberTypes.PublicProperties, typeof(KeyValuePair<,>))]
        [UnconditionalSuppressMessage(
            "Trimming",
            "IL2070",
            Justification = "The DynamicDependency roots KeyValuePair public properties; published NativeAOT scenarios verify Key and Value access for generic-only dictionary rows.")]
        private static DictionaryEntryAccessor CreateEntryAccessor(Type type) {
            return new DictionaryEntryAccessor(
                type.GetProperty("Key", BindingFlags.Instance | BindingFlags.Public),
                type.GetProperty("Value", BindingFlags.Instance | BindingFlags.Public));
        }

        private sealed class DictionaryEntryAccessor {
            internal DictionaryEntryAccessor(PropertyInfo? key, PropertyInfo? value) {
                Key = key;
                Value = value;
            }

            internal PropertyInfo? Key { get; }
            internal PropertyInfo? Value { get; }
            internal bool IsValid => Key?.CanRead == true && Value?.CanRead == true;
        }
    }

    internal readonly struct ObjectDictionaryEntry {
        internal ObjectDictionaryEntry(object? key, object? value) {
            Key = key;
            Value = value;
        }

        internal object? Key { get; }
        internal object? Value { get; }
    }
}
