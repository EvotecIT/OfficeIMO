using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeCffFontData {
    // Per-load caches keep shared Private DICT/Subr references shared after decoding.
    // The cumulative limits also cover distinct or overlapping metadata references.
    private sealed class CffParseContext {
        private readonly byte[] _data;
        private readonly int _end;
        private readonly Dictionary<(int Offset, int CountSize), (CffIndex Index, int Next)> _indices = new();
        private readonly Dictionary<(int Offset, int Length), CffDictionary> _dictionaries = new();
        private int _remainingIndexEntries = 2 * 1024 * 1024;
        private int _remainingMetadataBytes = 16 * 1024 * 1024;

        internal CffParseContext(byte[] data, int end) { _data = data; _end = end; }

        internal CffIndex ReadIndexAt(int offset, int countSize) => ReadIndex(ref offset, countSize);

        internal CffIndex ReadIndex(ref int cursor, int countSize) {
            var key = (cursor, countSize);
            if (_indices.TryGetValue(key, out var cached)) { cursor = cached.Next; return cached.Index; }
            var index = CffIndex.Read(_data, ref cursor, _end, countSize, ref _remainingIndexEntries, ref _remainingMetadataBytes);
            _indices.Add(key, (index, cursor));
            return index;
        }

        internal CffDictionary ReadDictionary(int offset, int length) {
            var key = (offset, length);
            if (_dictionaries.TryGetValue(key, out var cached)) return cached;
            EnsureRange(offset, length, 0, _end, "The CFF DICT range is invalid.");
            ChargeParseBudget(ref _remainingMetadataBytes, length);
            var dictionary = CffDictionary.Parse(_data, offset, length);
            _dictionaries.Add(key, dictionary);
            return dictionary;
        }
    }

    private static void ChargeParseBudget(ref int remaining, int count) {
        if (count < 0 || count > remaining) throw new InvalidDataException("CFF metadata exceeds the cumulative managed parsing limit.");
        remaining -= count;
    }
}
