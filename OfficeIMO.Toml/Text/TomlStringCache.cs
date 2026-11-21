namespace OfficeIMO.Toml;

/// <summary>
/// Reuses the strings of short repeated keys and values (<c>name</c>, <c>version</c>, <c>"https://pypi.org/simple"</c>)
/// instead of allocating each occurrence. Strings are immutable, so one cache per thread serves every reader and
/// document on it; a slot holds the last string that hashed there.
/// </summary>
internal static class TomlStringCache {
    private const int Size = 512;
    internal const int MaxLength = 32;

    [ThreadStatic] private static string?[]? t_entries;

    /// <summary>Returns a string equal to <paramref name="text"/>, reused when the same text was seen recently.</summary>
    public static string Get(ReadOnlySpan<char> text) {
        if (text.Length > MaxLength) return text.ToString();
        if (text.Length == 0) return string.Empty;
        string?[] entries = t_entries ??= new string?[Size];
        uint hash = (uint)text.Length;
        for (int i = 0; i < text.Length; i++) hash = unchecked(hash * 31 + text[i]);
        int slot = (int)((hash ^ (hash >> 15)) & (Size - 1));
        string? cached = entries[slot];
        if (cached != null && text.SequenceEqual(cached.AsSpan())) return cached;
        string created = text.ToString();
        entries[slot] = created;
        return created;
    }
}
