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
            private readonly OpenXmlPooledPartStream _stream;
            private readonly byte[] _bytes;
            private readonly List<AsciiTextEntry> _entries;
            private readonly object _textLock = new object();
            private string?[]? _text;

            internal IndexedAsciiItems(OpenXmlPooledPartStream stream, byte[] bytes, List<AsciiTextEntry> entries) {
                _stream = stream;
                _bytes = bytes;
                _entries = entries;
            }

            internal int Count => _entries.Count;

            internal string? Get(int index) {
                if ((uint)index >= (uint)_entries.Count) return null;
                string?[] text = Volatile.Read(ref _text) ?? InitializeText();
                string? value = Volatile.Read(ref text[index]);
                if (value != null) return value;
                AsciiTextEntry entry = _entries[index];
                value = Encoding.ASCII.GetString(_bytes, entry.Offset, entry.Length);
                return Interlocked.CompareExchange(ref text[index], value, null) ?? value;
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

            private string?[] InitializeText() {
                lock (_textLock) {
                    string?[]? text = Volatile.Read(ref _text);
                    if (text == null) {
                        text = new string?[_entries.Count];
                        Volatile.Write(ref _text, text);
                    }
                    return text;
                }
            }

            public void Dispose() => _stream.Dispose();
        }
    }
}
