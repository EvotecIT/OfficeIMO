using System.Threading;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    private const int MaximumCachedSubsetFiles = 512;
    private const long MaximumCachedSubsetBytes = 64L * 1024L * 1024L;
    private static readonly SubsetFontFileCache SubsetFileCache = new();

    // Content keys permit reuse across font-family snapshots without retaining the original font.
    internal byte[] BuildSubsetFontFile() => GetSubsetFontFileEntry().Subset;

    /// <summary>Returns the existing subset or its exact zlib payload, sharing one bounded cache.</summary>
    internal byte[] BuildSubsetFontFile(bool compress, out int uncompressedLength, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        SubsetCacheEntry entry = GetSubsetFontFileEntry();
        uncompressedLength = entry.Subset.Length;
        cancellationToken.ThrowIfCancellationRequested();
        return compress ? SubsetFileCache.GetCompressed(entry, cancellationToken) : entry.Subset;
    }

    private SubsetCacheEntry GetSubsetFontFileEntry() {
        IReadOnlyList<int> usedGlyphIds = GetUsedGlyphIds();
        int[] glyphIds = usedGlyphIds as int[] ?? System.Linq.Enumerable.ToArray(usedGlyphIds);
        var key = new SubsetCacheKey(_subsetFontFingerprint, glyphIds);
        if (SubsetFileCache.TryGet(key, out SubsetCacheEntry? cached)) {
            return cached!;
        }

        byte[] built = BuildSubsetFontFileUncached(glyphIds);
        return SubsetFileCache.AddOrGetExisting(key, built);
    }

    private readonly struct SubsetCacheKey : IEquatable<SubsetCacheKey> {
        private readonly SubsetFontFingerprint _fontFingerprint;
        private readonly int[] _glyphIds;
        private readonly int _hash;

        internal SubsetCacheKey(SubsetFontFingerprint fontFingerprint, int[] glyphIds) {
            _fontFingerprint = fontFingerprint;
            _glyphIds = glyphIds;
            int hash = fontFingerprint.GetHashCode();
            for (int i = 0; i < _glyphIds.Length; i++) {
                hash = (hash * 31) + _glyphIds[i];
            }

            _hash = hash;
        }

        public bool Equals(SubsetCacheKey other) {
            if (!_fontFingerprint.Equals(other._fontFingerprint) || _glyphIds.Length != other._glyphIds.Length) {
                return false;
            }

            for (int i = 0; i < _glyphIds.Length; i++) {
                if (_glyphIds[i] != other._glyphIds[i]) {
                    return false;
                }
            }

            return true;
        }

        public override bool Equals(object? obj) => obj is SubsetCacheKey other && Equals(other);
        public override int GetHashCode() => _hash;
    }

    private readonly struct SubsetFontFingerprint : IEquatable<SubsetFontFingerprint> {
        private readonly ulong _part0;
        private readonly ulong _part1;
        private readonly ulong _part2;
        private readonly ulong _part3;

        private SubsetFontFingerprint(byte[] hash) {
            _part0 = BitConverter.ToUInt64(hash, 0);
            _part1 = BitConverter.ToUInt64(hash, 8);
            _part2 = BitConverter.ToUInt64(hash, 16);
            _part3 = BitConverter.ToUInt64(hash, 24);
        }

        internal static SubsetFontFingerprint Create(byte[] fontData) {
#if NET5_0_OR_GREATER
            return new SubsetFontFingerprint(System.Security.Cryptography.SHA256.HashData(fontData));
#else
            using var sha256 = System.Security.Cryptography.SHA256.Create();
            return new SubsetFontFingerprint(sha256.ComputeHash(fontData));
#endif
        }

        public bool Equals(SubsetFontFingerprint other) =>
            _part0 == other._part0 &&
            _part1 == other._part1 &&
            _part2 == other._part2 &&
            _part3 == other._part3;

        public override bool Equals(object? obj) => obj is SubsetFontFingerprint other && Equals(other);

        public override int GetHashCode() {
            int hash = _part0.GetHashCode();
            hash = (hash * 397) ^ _part1.GetHashCode();
            hash = (hash * 397) ^ _part2.GetHashCode();
            return (hash * 397) ^ _part3.GetHashCode();
        }
    }

    private sealed class SubsetFontFileCache {
        private readonly object _syncRoot = new();
        private readonly Dictionary<SubsetCacheKey, LinkedListNode<SubsetCacheEntry>> _entries = new();
        private readonly LinkedList<SubsetCacheEntry> _recency = new();
        private long _cachedBytes;

        internal bool TryGet(SubsetCacheKey key, out SubsetCacheEntry? subset) {
            lock (_syncRoot) {
                if (_entries.TryGetValue(key, out LinkedListNode<SubsetCacheEntry>? node)) {
                    Touch(node);
                    subset = node.Value;
                    return true;
                }

                subset = null;
                return false;
            }
        }

        internal SubsetCacheEntry AddOrGetExisting(SubsetCacheKey key, byte[] subset) {
            lock (_syncRoot) {
                if (_entries.TryGetValue(key, out LinkedListNode<SubsetCacheEntry>? existing)) {
                    Touch(existing);
                    return existing.Value;
                }

                var entry = new SubsetCacheEntry(key, subset);
                if (subset.LongLength > MaximumCachedSubsetBytes) {
                    return entry;
                }

                while (_entries.Count >= MaximumCachedSubsetFiles ||
                       _cachedBytes + subset.LongLength > MaximumCachedSubsetBytes) {
                    EvictOldest();
                }

                LinkedListNode<SubsetCacheEntry> added = _recency.AddFirst(entry);
                _entries.Add(key, added);
                _cachedBytes += subset.LongLength;
                return entry;
            }
        }

        internal byte[] GetCompressed(SubsetCacheEntry entry, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            lock (_syncRoot) {
                cancellationToken.ThrowIfCancellationRequested();
                if (entry.Compressed != null) {
                    return entry.Compressed;
                }
            }

            // Compression is outside the shared lock: unrelated fonts and cache hits can continue.
            byte[] compressed = PdfFlateEncoder.Compress(entry.Subset, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            lock (_syncRoot) {
                cancellationToken.ThrowIfCancellationRequested();
                if (entry.Compressed != null) {
                    return entry.Compressed;
                }

                // An evicted or oversized subset may still be used by this document, but must not
                // retain extra bytes outside the cache's count and byte limits.
                if (!_entries.TryGetValue(entry.Key, out LinkedListNode<SubsetCacheEntry>? node) ||
                    !ReferenceEquals(node.Value, entry) ||
                    entry.Subset.LongLength + compressed.LongLength > MaximumCachedSubsetBytes) {
                    return compressed;
                }

                Touch(node);
                while (_cachedBytes + compressed.LongLength > MaximumCachedSubsetBytes) {
                    EvictOldest();
                }

                entry.Compressed = compressed;
                _cachedBytes += compressed.LongLength;
                return compressed;
            }
        }

        private void Touch(LinkedListNode<SubsetCacheEntry> node) {
            _recency.Remove(node);
            _recency.AddFirst(node);
        }

        private void EvictOldest() {
            LinkedListNode<SubsetCacheEntry> oldest = _recency.Last!;
            _recency.RemoveLast();
            _entries.Remove(oldest.Value.Key);
            _cachedBytes -= oldest.Value.Subset.LongLength + (oldest.Value.Compressed?.LongLength ?? 0);
        }
    }

    private sealed class SubsetCacheEntry {
        internal SubsetCacheEntry(SubsetCacheKey key, byte[] subset) {
            Key = key;
            Subset = subset;
        }

        internal SubsetCacheKey Key { get; }
        internal byte[] Subset { get; }
        // Read and published only under the owning cache lock.
        internal byte[]? Compressed { get; set; }
    }
}
