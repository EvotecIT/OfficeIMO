using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;

namespace PSWriteOffice.Services;

/// <summary>Formats nested collection and dictionary values consistently for Office table cells.</summary>
internal static class PowerShellCellValueFormatter
{
    public static bool TryFormatComplexValue(object? value, PowerShellObjectNormalizerOptions options, out string text)
    {
        if (!options.NormalizeCollectionValues)
        {
            text = string.Empty;
            return false;
        }

        value = Unwrap(value);
        if (value is string || value == null)
        {
            text = string.Empty;
            return false;
        }

        if (!PowerShellDictionaryAdapter.IsDictionaryLike(value) && value is not IEnumerable)
        {
            text = string.Empty;
            return false;
        }

        text = FormatValue(value, options, 0, new HashSet<object>(ReferenceComparer.Instance));
        return true;
    }

    private static string FormatValue(
        object? value,
        PowerShellObjectNormalizerOptions options,
        int depth,
        HashSet<object> activeObjects)
    {
        value = Unwrap(value);
        if (value == null)
        {
            return string.Empty;
        }

        if (value is string text)
        {
            return text;
        }

        if (PowerShellDictionaryAdapter.IsDictionaryLike(value))
        {
            EnsureContainerDepth(options, depth);
            if (!PowerShellDictionaryAdapter.TryGetEntries(value, options.MaxCollectionItems, out var entries))
            {
                return FormatScalar(value, options);
            }

            return TrackReference(value, activeObjects, () =>
            {
                var values = new List<string>(entries.Count);
                foreach (var entry in entries)
                {
                    values.Add(
                        FormatScalar(entry.Key, options) +
                        options.DictionaryKeyValueSeparator +
                        FormatValue(entry.Value, options, depth + 1, activeObjects));
                }

                return string.Join(options.DictionaryEntrySeparator, values);
            });
        }

        if (value is IEnumerable enumerable)
        {
            EnsureContainerDepth(options, depth);
            return TrackReference(value, activeObjects, () =>
            {
                var values = new List<string>();
                var enumerator = enumerable.GetEnumerator();
                try
                {
                    while (enumerator.MoveNext())
                    {
                        if (values.Count >= options.MaxCollectionItems)
                        {
                            throw new PowerShellNormalizationLimitException(
                                PowerShellNormalizationLimitMessages.Collection(
                                    "collection",
                                    options.MaxCollectionItems,
                                    values.Count + 1));
                        }

                        values.Add(FormatValue(enumerator.Current, options, depth + 1, activeObjects));
                    }
                }
                finally
                {
                    (enumerator as IDisposable)?.Dispose();
                }

                return string.Join(options.CollectionSeparator, values);
            });
        }

        return FormatScalar(value, options);
    }

    private static void EnsureContainerDepth(PowerShellObjectNormalizerOptions options, int depth)
    {
        if (depth >= options.MaxNestingDepth)
        {
            throw new PowerShellNormalizationLimitException(
                PowerShellNormalizationLimitMessages.Nesting(options.MaxNestingDepth, depth + 1));
        }
    }

    private static string TrackReference(object value, ISet<object> activeObjects, Func<string> formatter)
    {
        if (value.GetType().IsValueType)
        {
            return formatter();
        }

        if (!activeObjects.Add(value))
        {
            throw new InvalidDataException("The cell value contains a reference cycle.");
        }

        try
        {
            return formatter();
        }
        finally
        {
            activeObjects.Remove(value);
        }
    }

    private static object? Unwrap(object? value) => value is PSObject psObject ? psObject.BaseObject : value;

    private static string FormatScalar(object? value, PowerShellObjectNormalizerOptions options)
    {
        value = Unwrap(value);
        if (value == null)
        {
            return string.Empty;
        }

        if (value is bool boolValue)
        {
            return boolValue ? "True" : "False";
        }

        if (value is IFormattable formattable)
        {
            return formattable.ToString(null, options.Culture) ?? string.Empty;
        }

        return LanguagePrimitives.ConvertTo(value, typeof(string), options.Culture) as string ?? value.ToString() ?? string.Empty;
    }

    private sealed class ReferenceComparer : IEqualityComparer<object>
    {
        public static ReferenceComparer Instance { get; } = new();

        public new bool Equals(object? left, object? right) => ReferenceEquals(left, right);
        public int GetHashCode(object value) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
    }
}
