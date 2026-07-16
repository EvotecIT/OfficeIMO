using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>
    /// Retains a bounded prefix without separating the UTF-16 code units of a surrogate pair.
    /// A non-BMP Unicode scalar at a one-character boundary is retained intact.
    /// </summary>
    internal static string TruncateAdapterProjection(string value, int maxChars) {
        if (string.IsNullOrEmpty(value) || value.Length <= maxChars) return value;

        int length = Math.Max(1, maxChars);
        if (length < value.Length &&
            char.IsHighSurrogate(value[length - 1]) &&
            char.IsLowSurrogate(value[length])) {
            length = length == 1 ? 2 : length - 1;
        }

        return value.Substring(0, length);
    }

    /// <summary>
    /// Splits an adapter-owned text or Markdown projection into bounded pieces without
    /// discarding content or separating the UTF-16 code units of a surrogate pair.
    /// </summary>
    internal static IReadOnlyList<string> SplitAdapterProjection(string value, int maxChars) {
        return SplitAdapterProjection(value, maxChars, maxChars);
    }

    /// <summary>
    /// Splits an adapter-owned projection using separate first and subsequent part limits.
    /// A non-BMP Unicode scalar remains intact and may therefore occupy a two-code-unit part
    /// when the effective limit is one.
    /// </summary>
    internal static IReadOnlyList<string> SplitAdapterProjection(
        string value,
        int firstPartMaxChars,
        int subsequentPartMaxChars) {
        if (string.IsNullOrEmpty(value)) return Array.Empty<string>();

        int firstLimit = Math.Max(1, firstPartMaxChars);
        int subsequentLimit = Math.Max(1, subsequentPartMaxChars);
        if (value.Length <= firstLimit) return new[] { value };

        var parts = new List<string>(value.Length / subsequentLimit + 1);
        int offset = 0;
        int limit = firstLimit;
        while (offset < value.Length) {
            int length = Math.Min(limit, value.Length - offset);
            int boundary = offset + length;
            if (boundary < value.Length &&
                char.IsHighSurrogate(value[boundary - 1]) &&
                char.IsLowSurrogate(value[boundary])) {
                length = length == 1 ? 2 : length - 1;
            }

            parts.Add(value.Substring(offset, length));
            offset += length;
            limit = subsequentLimit;
        }
        return parts;
    }
}
