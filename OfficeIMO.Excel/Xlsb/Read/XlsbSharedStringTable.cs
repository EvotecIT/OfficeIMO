using OfficeIMO.Excel.Xlsb.Package;
using System.Buffers;
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
        private bool[]? _pooledTextPages;
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
                string?[]?[]? pages = _textPages;
                string?[] page = pages?[pageIndex] ?? RentTextPage(pageIndex);
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

        private string?[] RentTextPage(int pageIndex) {
            string?[]?[]? pages = _textPages;
            if (pages == null) {
                int pageCount = ((Count - 1) >> TextPageShift) + 1;
                pages = new string?[pageCount][];
                bool[] pooledPages = new bool[pageCount];
                _textPages = pages;
                _pooledTextPages = pooledPages;
            }

            int length = Math.Min(TextPageSize, Count - (pageIndex << TextPageShift));
            string?[] page = ArrayPool<string?>.Shared.Rent(length);
            bool pooled = page.Length <= TextPageSize;
            if (pooled) {
                // A shared rental may contain references left by an unrelated owner.
                Array.Clear(page, 0, page.Length);
            } else {
                // Keep retained reference storage bounded even if the pool changes its bucket policy.
                ArrayPool<string?>.Shared.Return(page, clearArray: true);
                page = new string?[length];
            }

            _pooledTextPages![pageIndex] = pooled;
            pages[pageIndex] = page;
            return page;
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
            XlsbPooledPartStream? part = _part;
            string?[]?[]? pages = _textPages;
            bool[]? pooledPages = _pooledTextPages;
            _part = null;
            _entries = null;
            _values = null;
            _textPages = null;
            _pooledTextPages = null;
            try {
                part?.Dispose();
            } finally {
                if (pages != null) {
                    for (int index = 0; index < pages.Length; index++) {
                        if (pooledPages![index] && pages[index] is string?[] page) {
                            ArrayPool<string?>.Shared.Return(page, clearArray: true);
                        }
                    }
                }
            }
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
