#nullable enable

using System.Buffers.Text;
using System.Globalization;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelXmlRangeDataReader {
            /// <summary>
            /// Reads a validated indexed number directly into the existing primitive
            /// cache. Date serials remain deferred so later object access materializes
            /// their DateTime value rather than the numeric projection.
            /// </summary>
            private bool TryGetUnloadedNumber(int ordinal, out double value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }

                value = default;
                if (_utf8Source == null || _currentRowIsBlank || _currentRow == null
                    || _currentValueLoaded[ordinal] || _owner._opt.NumericAsDecimal
                    || _owner._opt.CellValueConverter != null
                    || !_utf8Source.TryGetNumber(ordinal + _utf8SourceOrdinalOffset, out value, out bool dateStyle)) {
                    return false;
                }

                _currentPrimitiveKinds[ordinal] = XmlDataReaderPrimitiveKind.Double;
                _currentDoubleValues[ordinal] = value;
                _currentValues[ordinal] = null;
                _currentValueLoaded[ordinal] = !dateStyle;
                return true;
            }

            /// <summary>
            /// Reads an unloaded indexed date-style number without the general cell
            /// dispatch. The cached DateTime follows the existing getter-order rules;
            /// subsequent numeric access can still read the original serial.
            /// </summary>
            private bool TryGetUnloadedDateTime(int ordinal, out DateTime value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }

                value = default;
                if (_utf8Source == null || _currentRowIsBlank || _currentRow == null
                    || _currentValueLoaded[ordinal] || _owner._opt.NumericAsDecimal
                    || _owner._opt.CellValueConverter != null
                    || !_utf8Source.TryGetDateTime(ordinal + _utf8SourceOrdinalOffset, out value)) {
                    return false;
                }

                _currentPrimitiveKinds[ordinal] = XmlDataReaderPrimitiveKind.DateTime;
                _currentDateTimeValues[ordinal] = value;
                _currentValues[ordinal] = null;
                _currentValueLoaded[ordinal] = true;
                return true;
            }

            /// <summary>
            /// Reads an exact Int32 from an unloaded indexed number without entering
            /// the floating-point conversion pipeline. The shared cache still records
            /// a Double, preserving the value exposed by subsequent scalar getters.
            /// </summary>
            private bool TryGetUnloadedInt32(int ordinal, out int value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }

                value = default;
                if (_utf8Source == null || _currentRowIsBlank || _currentRow == null
                    || _currentValueLoaded[ordinal] || _owner._opt.NumericAsDecimal
                    || _owner._opt.CellValueConverter != null
                    || !_utf8Source.TryGetExactInt32(ordinal + _utf8SourceOrdinalOffset, out value)) {
                    return false;
                }

                _currentPrimitiveKinds[ordinal] = XmlDataReaderPrimitiveKind.Double;
                _currentDoubleValues[ordinal] = value;
                _currentValues[ordinal] = null;
                _currentValueLoaded[ordinal] = true;
                return true;
            }
        }

        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Parses a complete indexed numeric token through the common number
            /// parser. Formulas, missing values and XML-escaped text retain general
            /// value dispatch, which handles their additional semantics.
            /// </summary>
            internal bool TryGetNumber(int ordinal, out double value, out bool dateStyle) {
                EnsureNotDisposed();
                value = default;
                dateStyle = false;
                int cellIndex = _currentRowOffset + ordinal;
                byte encodedKind = _cellKinds![cellIndex];
                if ((encodedKind & CellKindMask) != (byte)Utf8CellKind.Number
                    || _formulaLengths != null && _formulaLengths[cellIndex] >= 0
                    || _valueLengths![cellIndex] < 0
                    || !TryParseDouble(TrimAsciiWhitespace(
                        _buffer!.AsSpan(_valueStarts![cellIndex], _valueLengths[cellIndex])), out value)) {
                    return false;
                }

                dateStyle = (encodedKind & DateStyleCellKindFlag) != 0;
                return true;
            }

            /// <summary>
            /// Converts only indexed date-style numeric cells through the ordinary
            /// workbook date-system conversion, including elapsed-time formats.
            /// </summary>
            internal bool TryGetDateTime(int ordinal, out DateTime value) {
                EnsureNotDisposed();
                value = default;
                byte encodedKind = _cellKinds![_currentRowOffset + ordinal];
                if ((encodedKind & DateStyleCellKindFlag) == 0
                    || !TryGetNumber(ordinal, out double number, out _)) {
                    return false;
                }

                value = _owner.FromExcelSerialDate(number, (encodedKind & CalendarStyleCellKindFlag) != 0);
                return true;
            }

            /// <summary>
            /// Reads only complete integer tokens within Int32 range. Date styles and
            /// formulas retain their existing conversion behavior. Negative zero also
            /// retains the floating-point path so later getters preserve its sign.
            /// </summary>
            internal bool TryGetExactInt32(int ordinal, out int value) {
                EnsureNotDisposed();
                value = default;
                int cellIndex = _currentRowOffset + ordinal;
                if (_cellKinds![cellIndex] != (byte)Utf8CellKind.Number
                    || _formulaLengths != null && _formulaLengths[cellIndex] >= 0
                    || _valueLengths![cellIndex] < 0) {
                    return false;
                }

                ReadOnlySpan<byte> token = TrimAsciiWhitespace(
                    _buffer!.AsSpan(_valueStarts![cellIndex], _valueLengths[cellIndex]));
                return Utf8Parser.TryParse(token, out value, out int consumed)
                    && consumed == token.Length
                    && (value != 0 || token[0] != (byte)'-');
            }
        }
    }
}
