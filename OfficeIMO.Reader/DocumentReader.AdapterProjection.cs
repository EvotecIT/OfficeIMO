using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>
    /// Splits an adapter-owned text or Markdown projection into fixed-size pieces that honor
    /// <see cref="ReaderOptions.MaxChars"/> without discarding content.
    /// </summary>
    internal static IReadOnlyList<string> SplitAdapterProjection(string value, int maxChars) {
        if (string.IsNullOrEmpty(value)) return Array.Empty<string>();

        int limit = Math.Max(1, maxChars);
        if (value.Length <= limit) return new[] { value };

        int capacity = value.Length / limit + (value.Length % limit == 0 ? 0 : 1);
        var parts = new List<string>(capacity);
        for (int offset = 0; offset < value.Length; offset += limit) {
            parts.Add(value.Substring(offset, Math.Min(limit, value.Length - offset)));
        }
        return parts;
    }
}
