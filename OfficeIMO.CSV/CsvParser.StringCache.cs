#if NET8_0_OR_GREATER
global using CsvStringCache = System.Collections.Generic.HashSet<string>;
#else
global using CsvStringCache = System.Collections.Generic.Dictionary<string, string>;
#endif

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    private static CsvStringCache? CreateStringCache(CsvParsingSettings options) =>
        options.InternStrings ? new CsvStringCache(StringComparer.Ordinal) : null;

    /// <summary>Shares normalized field strings only within the caller's read operation.</summary>
    private static string ReuseFieldString(string value, CsvStringCache? cache)
    {
        if (cache is null) return value;
        if (cache.TryGetValue(value, out string? cached)) return cached;
#if NET8_0_OR_GREATER
        cache.Add(value);
#else
        cache[value] = value;
#endif
        return value;
    }

#if NET8_0_OR_GREATER
    /// <summary>Avoids allocating a repeated, already normalized character field.</summary>
    private static string ReuseFieldSpan(ReadOnlySpan<char> value, CsvStringCache cache)
    {
#if NET9_0_OR_GREATER
        var lookup = cache.GetAlternateLookup<ReadOnlySpan<char>>();
        if (lookup.TryGetValue(value, out string? cached)) return cached;
        string owned = value.ToString();
        cache.Add(owned);
        return owned;
#else
        return ReuseFieldString(value.ToString(), cache);
#endif
    }
#endif
}
