using OfficeIMO.Excel.Xlsb.Package;
using System.Collections;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Owns a fully validated SST. Bounded tables retain their immutable UTF-16LE payload;
    /// canonical strings are cached only on demand. Reader references keep that payload
    /// alive when an internal workbook owner closes before its table readers.
    /// </summary>
    internal sealed partial class XlsbSharedStringTable : IReadOnlyList<string>, IDisposable {
        private const int MaximumIndexedPartBytes = 64 * 1024 * 1024;
        private const int MaximumIndexedItems = 1_048_576;
        private const int TextPageShift = 10;
        private const int TextPageSize = 1 << TextPageShift;
        private XlsbPooledPartStream? _part;
        private List<TextEntry>? _entries;
        private IReadOnlyList<string>? _values;
        private string?[]?[]? _textPages;
        private int _references = 1;
        private int _ownerReleased;

        private XlsbSharedStringTable(XlsbPooledPartStream part, List<TextEntry> entries) {
            _part = part;
            _entries = entries;
            Count = entries.Count;
        }

        private XlsbSharedStringTable(IReadOnlyList<string> values) {
            _values = values;
            Count = values.Count;
        }

        public int Count { get; }

        public string this[int index] {
            get {
                if ((uint)index >= (uint)Count) throw new ArgumentOutOfRangeException(nameof(index));
                if (_values != null) return _values[index];
                List<TextEntry> entries = _entries ?? throw new ObjectDisposedException(nameof(XlsbSharedStringTable));
                int pageIndex = index >> TextPageShift;
                string?[]?[] pages = _textPages ??= new string?[((Count - 1) >> TextPageShift) + 1][];
                string?[] page = pages[pageIndex] ??= new string?[Math.Min(TextPageSize, Count - (pageIndex << TextPageShift))];
                int pageOffset = index & (TextPageSize - 1);
                string? value = page[pageOffset];
                if (value == null) {
                    TextEntry entry = entries[index];
                    value = Encoding.Unicode.GetString(_part!.Buffer, entry.Offset, entry.ByteCount);
                    page[pageOffset] = value;
                }
                return value;
            }
        }

        /// <summary>Provides validated little-endian code units without materializing canonical text.</summary>
        internal bool TryGetUtf16(int index, out ArraySegment<byte> text) {
            if (_entries == null) {
                text = default;
                return false;
            }
            TextEntry entry = _entries[index];
            text = new ArraySegment<byte>(_part!.Buffer, entry.Offset, entry.ByteCount);
            return true;
        }

        internal void AddReaderReference() {
            if (Volatile.Read(ref _ownerReleased) != 0) throw new ObjectDisposedException(nameof(XlsbSharedStringTable));
            int references;
            do {
                references = Volatile.Read(ref _references);
                if (references == 0) throw new ObjectDisposedException(nameof(XlsbSharedStringTable));
            } while (Interlocked.CompareExchange(ref _references, references + 1, references) != references);
        }

        internal void ReleaseReaderReference() {
            if (Interlocked.Decrement(ref _references) != 0) return;
            _part?.Dispose();
            _part = null;
            _entries = null;
            _values = null;
            _textPages = null;
        }

        public void Dispose() {
            if (Interlocked.Exchange(ref _ownerReleased, 1) == 0) ReleaseReaderReference();
        }

        public IEnumerator<string> GetEnumerator() {
            for (int index = 0; index < Count; index++) yield return this[index];
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        private readonly struct TextEntry {
            internal TextEntry(int offset, int byteCount) { Offset = offset; ByteCount = byteCount; }
            internal int Offset { get; }
            internal int ByteCount { get; }
        }
    }
}
