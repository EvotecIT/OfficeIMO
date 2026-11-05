namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    private static Dictionary<string, string>? CreateStringCache(CsvParsingSettings options) =>
        options.InternStrings ? new Dictionary<string, string>(StringComparer.Ordinal) : null;

    /// <summary>Shares normalized field strings only within the caller's read operation.</summary>
    private static string ReuseFieldString(string value, Dictionary<string, string>? cache)
    {
        if (cache is null) return value;
#if NET8_0_OR_GREATER
        ref string? entry = ref System.Runtime.InteropServices.CollectionsMarshal.GetValueRefOrAddDefault(
            cache, value, out bool exists);
        if (!exists) entry = value;
        return entry!;
#else
        if (cache.TryGetValue(value, out string? cached)) return cached;
        cache[value] = value;
        return value;
#endif
    }

#if NET8_0_OR_GREATER
    /// <summary>Avoids allocating a repeated, already normalized character field.</summary>
    private static string ReuseFieldSpan(ReadOnlySpan<char> value, Dictionary<string, string> cache)
    {
#if NET9_0_OR_GREATER
        var lookup = cache.GetAlternateLookup<ReadOnlySpan<char>>();
        if (lookup.TryGetValue(value, out string? cached)) return cached;
        string owned = value.ToString();
        cache.Add(owned, owned);
        return owned;
#else
        return ReuseFieldString(value.ToString(), cache);
#endif
    }
#endif
}
