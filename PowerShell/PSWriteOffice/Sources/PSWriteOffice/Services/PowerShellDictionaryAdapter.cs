using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Reflection;

namespace PSWriteOffice.Services;

/// <summary>Adapts PowerShell, non-generic, and generic dictionary implementations.</summary>
internal static class PowerShellDictionaryAdapter
{
    private static readonly ConcurrentDictionary<Type, bool> DictionaryTypeCache = new();
    private static readonly ConcurrentDictionary<Type, DictionaryEntryAccessor> EntryAccessorCache = new();

    public static bool IsDictionaryLike(object? value)
    {
        if (value is PSObject psObject)
        {
            value = psObject.BaseObject;
        }

        return value is IDictionary ||
            value != null && DictionaryTypeCache.GetOrAdd(value.GetType(), IsGenericDictionaryType);
    }

    public static bool TryGetEntries(object? value, int maximumItems, out IReadOnlyList<PowerShellDictionaryEntry> entries)
    {
        var result = new List<PowerShellDictionaryEntry>();
        entries = result;
        if (value == null)
        {
            return false;
        }

        if (maximumItems <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(maximumItems));
        }

        if (value is PSObject psObject)
        {
            value = psObject.BaseObject;
        }

        if (value is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                AddBounded(result, new PowerShellDictionaryEntry(entry.Key, entry.Value), maximumItems);
            }

            return true;
        }

        var type = value.GetType();
        if (!DictionaryTypeCache.GetOrAdd(type, IsGenericDictionaryType) || value is not IEnumerable enumerable)
        {
            return false;
        }

        foreach (var item in enumerable)
        {
            if (item == null)
            {
                continue;
            }

            var accessor = EntryAccessorCache.GetOrAdd(item.GetType(), CreateEntryAccessor);
            if (!accessor.IsValid)
            {
                result.Clear();
                return false;
            }

            AddBounded(
                result,
                new PowerShellDictionaryEntry(accessor.Key!.GetValue(item), accessor.Value!.GetValue(item)),
                maximumItems);
        }

        return true;
    }

    /// <summary>
    /// Enumerates a dictionary used as the top-level row schema. Nested-cell limits do not apply here;
    /// the consuming table or file format owns its schema and column boundaries.
    /// </summary>
    public static bool TryGetRowEntries(object? value, out IReadOnlyList<PowerShellDictionaryEntry> entries)
        => TryGetEntries(value, int.MaxValue, out entries);

    public static object? GetValue(
        IReadOnlyList<PowerShellDictionaryEntry> entries,
        string column,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        foreach (var entry in entries)
        {
            if (string.Equals(entry.Key?.ToString(), column, comparison))
            {
                return entry.Value;
            }
        }

        return null;
    }

    private static void AddBounded(
        ICollection<PowerShellDictionaryEntry> entries,
        PowerShellDictionaryEntry entry,
        int maximumItems)
    {
        if (entries.Count >= maximumItems)
        {
            throw new PowerShellNormalizationLimitException(
                PowerShellNormalizationLimitMessages.Collection(
                    "dictionary",
                    maximumItems,
                    entries.Count + 1));
        }

        entries.Add(entry);
    }

    private static bool IsGenericDictionaryType(Type type)
    {
        return type.GetInterfaces().Any(interfaceType =>
        {
            if (!interfaceType.IsGenericType)
            {
                return false;
            }

            var definition = interfaceType.GetGenericTypeDefinition();
            return definition == typeof(IDictionary<,>) || definition == typeof(IReadOnlyDictionary<,>);
        });
    }

    private static DictionaryEntryAccessor CreateEntryAccessor(Type type)
    {
        return new DictionaryEntryAccessor(
            type.GetProperty("Key", BindingFlags.Instance | BindingFlags.Public),
            type.GetProperty("Value", BindingFlags.Instance | BindingFlags.Public));
    }

    private sealed class DictionaryEntryAccessor
    {
        public DictionaryEntryAccessor(PropertyInfo? key, PropertyInfo? value)
        {
            Key = key;
            Value = value;
        }

        public PropertyInfo? Key { get; }
        public PropertyInfo? Value { get; }
        public bool IsValid => Key?.CanRead == true && Value?.CanRead == true;
    }
}

internal readonly struct PowerShellDictionaryEntry
{
    public PowerShellDictionaryEntry(object? key, object? value)
    {
        Key = key;
        Value = value;
    }

    public object? Key { get; }
    public object? Value { get; }
}
