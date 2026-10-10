using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class SharedStringCache {
        private readonly struct AsciiTextEntry {
            internal AsciiTextEntry(int offset, int length) {
                Offset = offset;
                Length = length;
            }

            internal int Offset { get; }
            internal int Length { get; }
        }

        /// <summary>
        /// Owns a fully validated table buffer. Only requested UTF-16 strings are decoded;
        /// borrowed UTF-8 text uses the canonical ASCII bytes until the workbook closes.
        /// </summary>
        private sealed class IndexedAsciiItems : IDisposable {
            private const int TextPageShift = 10;
            private const int TextPageSize = 1 << TextPageShift;
            private readonly OpenXmlPooledPartStream _stream;
            private readonly byte[] _bytes;
            private readonly List<AsciiTextEntry> _entries;
            private readonly object _textLock = new object();
            private string?[]?[]? _textPages;
            private int _disposed;

            internal IndexedAsciiItems(OpenXmlPooledPartStream stream, byte[] bytes, List<AsciiTextEntry> entries) {
                _stream = stream;
                _bytes = bytes;
                _entries = entries;
            }

            internal int Count => _entries.Count;

            internal string? Get(int index) {
                if ((uint)index >= (uint)_entries.Count) return null;
                int pageIndex = index >> TextPageShift;
                string?[]?[]? pages = Volatile.Read(ref _textPages);
                string?[] text = (pages == null ? null : Volatile.Read(ref pages[pageIndex]))
                    ?? InitializeTextPage(pageIndex);
                int pageOffset = index & (TextPageSize - 1);
                string? value = Volatile.Read(ref text[pageOffset]);
                if (value != null) return value;
                AsciiTextEntry entry = _entries[index];
                value = Encoding.ASCII.GetString(_bytes, entry.Offset, entry.Length);
                return Interlocked.CompareExchange(ref text[pageOffset], value, null) ?? value;
            }

            internal bool TryGetUtf8(int index, out ArraySegment<byte> value) {
                if ((uint)index >= (uint)_entries.Count) {
                    value = default;
                    return false;
                }
                AsciiTextEntry entry = _entries[index];
                value = new ArraySegment<byte>(_bytes, entry.Offset, entry.Length);
                return true;
            }

            // Materializing a few strings should not allocate slots for the complete table.
            private string?[] InitializeTextPage(int pageIndex) {
                lock (_textLock) {
                    string?[]?[]? pages = Volatile.Read(ref _textPages);
                    if (pages == null) {
                        int pageCount = ((_entries.Count - 1) >> TextPageShift) + 1;
                        pages = new string?[pageCount][];
                        Volatile.Write(ref _textPages, pages);
                    }
                    string?[]? text = Volatile.Read(ref pages[pageIndex]);
                    if (text == null) {
                        int pageLength = Math.Min(TextPageSize, _entries.Count - (pageIndex << TextPageShift));
                        text = new string?[pageLength];
                        Volatile.Write(ref pages[pageIndex], text);
                    }
                    return text;
                }
            }

            public void Dispose() {
                if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
                try {
                    _stream.Dispose();
                } finally {
                    IndexedAsciiEntryPool.Return(_entries);
                }
            }
        }
    }
}
