using System;

namespace OfficeIMO.Core.Internal;

internal static class OfficeDocumentHeadingPath {
    internal const int MaximumCharacters = 1_024;

    internal static string Append(string? parent, string? value, string separator) {
        if (separator == null) throw new ArgumentNullException(nameof(separator));
        parent ??= string.Empty;
        value ??= string.Empty;
        if (parent.Length == 0) return Truncate(value);
        if (value.Length == 0) return Truncate(parent);
        if (separator.Length >= MaximumCharacters || value.Length >= MaximumCharacters - separator.Length) return Truncate(value);

        int prefixBudget = MaximumCharacters - separator.Length - value.Length;
        if (parent.Length <= prefixBudget) return parent + separator + value;
        return parent.Substring(0, SafePrefixLength(parent, prefixBudget - 1)) + "…" + separator + value;
    }

    private static string Truncate(string value) => value.Length <= MaximumCharacters
        ? value
        : value.Substring(0, SafePrefixLength(value, MaximumCharacters));

    private static int SafePrefixLength(string value, int maximum) {
        int length = Math.Min(value.Length, maximum);
        if (length > 0 && length < value.Length && char.IsHighSurrogate(value[length - 1]) && char.IsLowSurrogate(value[length])) length--;
        return length;
    }
}
