#nullable enable

#if NET8_0_OR_GREATER

using System.Globalization;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelXmlRangeDataReader : IDataReaderUtf8TextSource {
            public bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }
                if (!_currentRowIsBlank && _utf8Source != null) {
                    return _utf8Source.TryGetUtf8Text(ordinal + _utf8SourceOrdinalOffset, out text);
                }

                text = default;
                return false;
            }
        }

        private sealed partial class ExcelRangeDataReader : IDataReaderUtf8TextSource {
            public bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }

                text = default;
                return false;
            }
        }

        private sealed partial class ExcelUtf8RangeRowSource {
            internal bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text) {
                EnsureNotDisposed();
                int cellIndex = _currentRowOffset + ordinal;
                Utf8CellKind kind = (Utf8CellKind)(_cellKinds![cellIndex] & CellKindMask);
                if ((kind != Utf8CellKind.InlineString && kind != Utf8CellKind.String && kind != Utf8CellKind.SharedString)
                    || (_formulaLengths != null && _formulaLengths[cellIndex] >= 0)
                    || !TryGetUtf8Value(ordinal, out ArraySegment<byte> value)) {
                    text = default;
                    return false;
                }

                text = value.AsSpan();
                return true;
            }
        }
    }
}
#endif
