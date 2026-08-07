using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace OfficeIMO.Data {
    /// <summary>
    /// Adapts non-generic and generic dictionary implementations without making
    /// <see cref="ObjectFlattener"/> depend on any one dictionary concrete type.
    /// </summary>
    internal static class ObjectDictionaryAdapter {
        private static readonly ConcurrentDictionary<Type, bool> DictionaryTypeCache = new();
        private static readonly ConcurrentDictionary<Type, DictionaryEntryAccessor> EntryAccessorCache = new();

        internal static bool TryGetEntries(object? value, int maximumItems, out List<ObjectDictionaryEntry> entries) {
            entries = new List<ObjectDictionaryEntry>();
            if (value == null) return false;
            if (maximumItems <= 0) throw new ArgumentOutOfRangeException(nameof(maximumItems));

            if (value is IDictionary dictionary) {
                foreach (DictionaryEntry entry in dictionary) {
                    AddBounded(entries, new ObjectDictionaryEntry(entry.Key, entry.Value), maximumItems);
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

                AddBounded(
                    entries,
                    new ObjectDictionaryEntry(accessor.Key!.GetValue(item), accessor.Value!.GetValue(item)),
                    maximumItems);
            }

            return true;
        }

        private static void AddBounded(List<ObjectDictionaryEntry> entries, ObjectDictionaryEntry entry, int maximumItems) {
            if (entries.Count >= maximumItems) {
                throw new InvalidDataException($"The dictionary exceeds the {maximumItems}-item flattening limit.");
            }
            entries.Add(entry);
        }

        private static bool IsGenericDictionaryType(Type type) {
            return type.GetInterfaces().Any(interfaceType => {
                if (!interfaceType.IsGenericType) return false;
                Type definition = interfaceType.GetGenericTypeDefinition();
                return definition == typeof(IDictionary<,>) || definition == typeof(IReadOnlyDictionary<,>);
            });
        }

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
