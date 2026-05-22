using System;
using System.Globalization;

namespace OfficeIMO.Excel {
    internal static class InvariantNumberText {
        private const int InitialCacheSize = 4096;
        private const int MaxCacheSize = 65536;
        private static readonly object CacheLock = new object();
        private static string[] NonNegativeCache = CreateNonNegativeCache(InitialCacheSize);

        internal static string Get(int value) {
            var cache = NonNegativeCache;
            return (uint)value < (uint)cache.Length
                ? cache[value]
                : (uint)value < MaxCacheSize
                    ? GetOrAddNonNegative(value)
                : value.ToString(CultureInfo.InvariantCulture);
        }

        internal static string Get(uint value) {
            var cache = NonNegativeCache;
            return value < (uint)cache.Length
                ? cache[value]
                : value < MaxCacheSize
                    ? GetOrAddNonNegative((int)value)
                : value.ToString(CultureInfo.InvariantCulture);
        }

        internal static string Get(long value) {
            var cache = NonNegativeCache;
            return (ulong)value < (ulong)cache.Length
                ? cache[(int)value]
                : (ulong)value < MaxCacheSize
                    ? GetOrAddNonNegative((int)value)
                : value.ToString(CultureInfo.InvariantCulture);
        }

        internal static string Get(ulong value) {
            var cache = NonNegativeCache;
            return value < (ulong)cache.Length
                ? cache[(int)value]
                : value < MaxCacheSize
                    ? GetOrAddNonNegative((int)value)
                : value.ToString(CultureInfo.InvariantCulture);
        }

        internal static bool TryGet(int value, out string text) {
            var cache = NonNegativeCache;
            if ((uint)value < (uint)cache.Length) {
                text = cache[value];
                return true;
            }

            text = string.Empty;
            return false;
        }

        internal static bool TryGet(long value, out string text) {
            var cache = NonNegativeCache;
            if ((ulong)value < (ulong)cache.Length) {
                text = cache[(int)value];
                return true;
            }

            text = string.Empty;
            return false;
        }

        internal static bool TryGet(ulong value, out string text) {
            var cache = NonNegativeCache;
            if (value < (ulong)cache.Length) {
                text = cache[(int)value];
                return true;
            }

            text = string.Empty;
            return false;
        }

        private static string GetOrAddNonNegative(int value) {
            lock (CacheLock) {
                var cache = NonNegativeCache;
                if ((uint)value < (uint)cache.Length) {
                    return cache[value];
                }

                int expandedSize = cache.Length;
                do {
                    expandedSize *= 2;
                } while ((uint)value >= (uint)expandedSize && expandedSize < MaxCacheSize);

                if (expandedSize > MaxCacheSize) {
                    expandedSize = MaxCacheSize;
                }

                var expanded = new string[expandedSize];
                Array.Copy(cache, expanded, cache.Length);
                for (int i = cache.Length; i < expanded.Length; i++) {
                    expanded[i] = i.ToString(CultureInfo.InvariantCulture);
                }

                NonNegativeCache = expanded;
                return expanded[value];
            }
        }

        private static string[] CreateNonNegativeCache(int cacheSize) {
            var cache = new string[cacheSize];
            for (int i = 0; i < cache.Length; i++) {
                cache[i] = i.ToString(CultureInfo.InvariantCulture);
            }

            return cache;
        }
    }
}
