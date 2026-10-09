using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>Retains a bounded set of immutable UTF-8 encodings of canonical reader text.</summary>
    internal sealed class ExcelUtf8TextCache {
        internal const int SlotCount = 256;
        internal const int MaximumCachedItemBytes = 4 * 1024;
        private Entry?[]? _entries;

        /// <summary>
        /// Encodes text on a miss. Keys may be reused for another row; string identity keeps those
        /// entries distinct. Returned arrays are never overwritten, including after slot eviction.
        /// Values above the retention limit are encoded without retaining their array in the cache.
        /// </summary>
        internal ArraySegment<byte> Get(int key, string text) {
            Entry?[] entries = Volatile.Read(ref _entries) ?? InitializeEntries();
            int slot = key & (SlotCount - 1);
            Entry? entry = Volatile.Read(ref entries[slot]);
            if (entry != null && entry.Key == key && ReferenceEquals(entry.Text, text)) {
                return new ArraySegment<byte>(entry.Bytes);
            }

            byte[] bytes = Encoding.UTF8.GetBytes(text);
            if (bytes.Length <= MaximumCachedItemBytes) {
                entry = Volatile.Read(ref entries[slot]);
                if (entry != null && entry.Key == key && ReferenceEquals(entry.Text, text)) {
                    bytes = entry.Bytes;
                } else {
                    Volatile.Write(ref entries[slot], new Entry(key, text, bytes));
                }
            }

            return new ArraySegment<byte>(bytes);
        }

        private Entry?[] InitializeEntries() {
            var entries = new Entry?[SlotCount];
            return Interlocked.CompareExchange(ref _entries, entries, null) ?? entries;
        }

        private sealed class Entry {
            internal Entry(int key, string text, byte[] bytes) {
                Key = key;
                Text = text;
                Bytes = bytes;
            }

            internal int Key { get; }
            internal string Text { get; }
            internal byte[] Bytes { get; }
        }
    }
}
