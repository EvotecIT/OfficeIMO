#if NET8_0_OR_GREATER

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    // One empty string set; its buckets and entries occupy less than 7 MiB on a 64-bit runtime.
    private const int MaximumRetainedStringCacheCapacity = 327680;
    private static CsvStringCache? _availableStringCache;

    /// <summary>Leases storage exclusively while preserving a separate string cache for every read.</summary>
    private static CsvStringCache? RentStringCacheWorkspace(CsvParsingSettings options)
    {
        if (!options.InternStrings) return null;
        return Interlocked.Exchange(ref _availableStringCache, null) ?? CreateStringCache(options);
    }

    /// <summary>Reuses bounded storage only after every source has dropped its alias and all strings are cleared.</summary>
    private static void ReturnStringCacheWorkspace(CsvStringCache? cache)
    {
        if (cache is null) return;
        int capacity = cache.EnsureCapacity(0);
        // A small read must not repeatedly clear the buckets left by a much larger read.
        if (capacity > MaximumRetainedStringCacheCapacity ||
            (capacity > 1024 && cache.Count < capacity / 4)) return;
        cache.Clear();
        Interlocked.CompareExchange(ref _availableStringCache, cache, null);
    }
}
#endif
