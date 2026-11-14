using System.Collections.Concurrent;

namespace OfficeIMO.Drawing;

/// <summary>Shares the four-character strings of OpenType layout tags read from font tables.</summary>
/// <remarks>Shaping reads feature and script tags for every run; fonts use a small set of them.</remarks>
internal static class OfficeOpenTypeTags {
    // Hostile fonts can name many distinct tags; past this bound tags are allocated as before.
    private const int MaximumSharedTags = 2048;
    private static readonly ConcurrentDictionary<uint, string> Tags = new();

    internal static string ToTag(uint value) {
        if (Tags.TryGetValue(value, out string? tag)) return tag;
        tag = new string(new[] {
            (char)((value >> 24) & 0xFF),
            (char)((value >> 16) & 0xFF),
            (char)((value >> 8) & 0xFF),
            (char)(value & 0xFF)
        });
        return Tags.Count < MaximumSharedTags ? Tags.GetOrAdd(value, tag) : tag;
    }
}
