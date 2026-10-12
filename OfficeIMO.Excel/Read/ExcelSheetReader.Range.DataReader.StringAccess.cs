#nullable enable

using System.Globalization;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelXmlRangeDataReader {
            /// <summary>
            /// Loads qualified indexed text into the ordinary current-value cache
            /// without dispatching through numeric, date, or formula conversions.
            /// </summary>
            private bool TryGetUnloadedString(int ordinal, out string value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }

                value = string.Empty;
                if (_utf8Source == null || _currentRowIsBlank || _currentRow == null
                    || _currentValueLoaded[ordinal] || _owner._opt.CellValueConverter != null
                    || !_utf8Source.TryGetString(ordinal + _utf8SourceOrdinalOffset, out value)) {
                    return false;
                }

                _currentPrimitiveKinds[ordinal] = XmlDataReaderPrimitiveKind.None;
                _currentValues[ordinal] = value;
                _currentValueLoaded[ordinal] = true;
                return true;
            }
        }

        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Reads only qualified text cells using the existing XML normalization
            /// and shared-string decoding. Other cell kinds retain value dispatch.
            /// </summary>
            internal bool TryGetString(int ordinal, out string value) {
                EnsureNotDisposed();
                value = string.Empty;
                int cellIndex = _currentRowOffset + ordinal;
                Utf8CellKind kind = (Utf8CellKind)(_cellKinds![cellIndex] & CellKindMask);
                if ((kind != Utf8CellKind.String && kind != Utf8CellKind.InlineString
                    && kind != Utf8CellKind.SharedString)
                    || _formulaLengths != null && _formulaLengths[cellIndex] >= 0) {
                    return false;
                }

                int length = _valueLengths![cellIndex];
                if (kind == Utf8CellKind.SharedString) {
                    if (length != SharedStringIndexValueLength) return false;
                    string? sharedText = _owner.GetSharedString(_valueStarts![cellIndex]);
                    if (sharedText == null) return false;
                    value = sharedText;
                    return true;
                }

                if (length < 0) return false;
                value = DecodeString(_valueStarts![cellIndex], length);
                return true;
            }
        }
    }
}
