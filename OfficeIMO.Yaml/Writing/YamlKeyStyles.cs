using System.Runtime.CompilerServices;

namespace OfficeIMO.Yaml;

/// <summary>
/// The default style of a block mapping key, remembered per string instance. Dictionaries write the same key strings
/// over and over (a reader hands out one instance per repeated key, and code uses literals), so the scans behind the
/// choice run once per key instead of once per entry. The choice depends only on the text, so a remembered one never
/// goes stale. Each thread keeps a small direct-mapped table; a collision replaces the slot, and only short keys are
/// kept, so the memory it holds stays bounded.
/// </summary>
internal static class YamlKeyStyles {
    private const int Size = 256;
    private const int MaxLength = 128;

    private struct Entry {
        public string? Key;
        public YamlStyle Style;
    }

    [ThreadStatic] private static Entry[]? t_entries;

    /// <summary>The style <see cref="YamlScalarFormat.Choose"/> picks for the text as a key of a block mapping.</summary>
    public static YamlStyle Block(string key) {
        if (key.Length > MaxLength) return Choose(key);
        Entry[] entries = t_entries ??= new Entry[Size];
        ref Entry entry = ref entries[RuntimeHelpers.GetHashCode(key) & (Size - 1)];
        if (ReferenceEquals(entry.Key, key)) return entry.Style;
        YamlStyle style = Choose(key);
        entry.Key = key;
        entry.Style = style;
        return style;
    }

    /// <summary>Whether the text is written plain on one line as a key of a block mapping, as an implicit key.</summary>
    public static bool IsPlain(string key) => key.Length <= YamlScalarFormat.MaxImplicitKeyLength && Block(key) == YamlStyle.Plain;

    private static YamlStyle Choose(string key) => YamlScalarFormat.Choose(key, YamlStyle.None, flow: false, key: true, YamlStyle.SingleQuoted);
}
