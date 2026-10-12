#nullable enable

using System.Buffers;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private Utf8ImplicitCellTagCacheEntry[]? _implicitCellTagCache;

            /// <summary>
            /// Holds one validated implicit start tag per projected column. Entries
            /// contain offsets into this source's immutable worksheet buffer, never
            /// coordinates or values, so a provisional row rollback can retain them.
            /// </summary>
            private Utf8ImplicitCellTagCacheEntry[] RentImplicitCellTagCache() {
                var cache = ArrayPool<Utf8ImplicitCellTagCacheEntry>.Shared.Rent(_fieldCount);
                Array.Clear(cache, 0, _fieldCount);
                _implicitCellTagCache = cache;
                return cache;
            }

            private void ReturnImplicitCellTagCache() {
                if (_implicitCellTagCache != null) {
                    ArrayPool<Utf8ImplicitCellTagCacheEntry>.Shared.Return(_implicitCellTagCache);
                    _implicitCellTagCache = null;
                }
            }

            private readonly struct Utf8ImplicitCellTagCacheEntry {
                internal Utf8ImplicitCellTagCacheEntry(
                    int start,
                    int length,
                    Utf8CellKind kind,
                    int styleIndex,
                    bool isEmpty,
                    byte encodedKind,
                    bool dateStylesEnabled) {
                    Start = start;
                    Length = length;
                    Kind = kind;
                    StyleIndex = styleIndex;
                    IsEmpty = isEmpty;
                    EncodedKind = encodedKind;
                    DateStylesEnabled = dateStylesEnabled;
                }

                internal int Start { get; }
                internal int Length { get; }
                internal int StyleIndex { get; }
                internal Utf8CellKind Kind { get; }
                internal byte EncodedKind { get; }
                internal bool IsEmpty { get; }
                internal bool DateStylesEnabled { get; }
            }
        }
    }
}
