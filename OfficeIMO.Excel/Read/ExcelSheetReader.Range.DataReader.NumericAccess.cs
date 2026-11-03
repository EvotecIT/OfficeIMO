using System.Buffers.Text;
using System.Globalization;
using System.Runtime.CompilerServices;
#if NET8_0_OR_GREATER
using System.Runtime.InteropServices;
#endif

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelXmlRangeDataReader {
            /// <summary>
            /// Accesses a cache slot after the caller's open-row and ordinal checks.
            /// Every cache array is readonly and allocated with the fixed field
            /// count in the constructor. The ordinal has already been checked
            /// against that same count before reaching this helper.
            /// Legacy runtimes retain ordinary array indexing.
            /// </summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static ref T GetValidatedOrdinalSlot<T>(T[] values, int ordinal) {
#if NET8_0_OR_GREATER
                return ref Unsafe.Add(ref MemoryMarshal.GetArrayDataReference(values), ordinal);
#else
                return ref values[ordinal];
#endif
            }

            /// <summary>
            /// Reads a validated indexed number directly into the existing primitive
            /// cache. Successful unloaded primitives have no object reference to clear.
            /// Date serials remain deferred so later object access materializes
            /// their DateTime value rather than the numeric projection.
            /// </summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private bool TryGetUnloadedNumber(int ordinal, out double value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }

                value = default;
                if (_utf8Source == null || _currentRowIsBlank || _currentRow == null) return false;
                ref bool loaded = ref _currentValueLoaded[ordinal];
                if (loaded || _owner._opt.NumericAsDecimal
                    || _owner._opt.CellValueConverter != null
                    || !_utf8Source.TryGetNumber(ordinal + _utf8SourceOrdinalOffset, out value, out bool dateStyle)) {
                    return false;
                }

                GetValidatedOrdinalSlot(_currentPrimitiveKinds, ordinal) = XmlDataReaderPrimitiveKind.Double;
                GetValidatedOrdinalSlot(_currentDoubleValues, ordinal) = value;
                loaded = !dateStyle;
                return true;
            }

            /// <summary>
            /// Reads an unloaded indexed date-style number without the general cell
            /// dispatch. The cached DateTime follows the existing getter-order rules;
            /// subsequent numeric access can still read the original serial.
            /// </summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private bool TryGetUnloadedDateTime(int ordinal, out DateTime value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }

                value = default;
                if (_utf8Source == null || _currentRowIsBlank || _currentRow == null) return false;
                ref bool loaded = ref _currentValueLoaded[ordinal];
                if (loaded || _owner._opt.NumericAsDecimal
                    || _owner._opt.CellValueConverter != null
                    || !_utf8Source.TryGetDateTime(ordinal + _utf8SourceOrdinalOffset, out value)) {
                    return false;
                }

                GetValidatedOrdinalSlot(_currentPrimitiveKinds, ordinal) = XmlDataReaderPrimitiveKind.DateTime;
                GetValidatedOrdinalSlot(_currentDateTimeValues, ordinal) = value;
                loaded = true;
                return true;
            }

            /// <summary>
            /// Reads an exact Int32 from an unloaded indexed number without entering
            /// the floating-point conversion pipeline. The shared cache still records
            /// a Double, preserving the value exposed by subsequent scalar getters.
            /// </summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private bool TryGetUnloadedInt32(int ordinal, out int value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }

                value = default;
                if (_utf8Source == null || _currentRowIsBlank || _currentRow == null) return false;
                ref bool loaded = ref _currentValueLoaded[ordinal];
                if (loaded || _owner._opt.NumericAsDecimal
                    || _owner._opt.CellValueConverter != null
                    || !_utf8Source.TryGetExactInt32(ordinal + _utf8SourceOrdinalOffset, out value)) {
                    return false;
                }

                GetValidatedOrdinalSlot(_currentPrimitiveKinds, ordinal) = XmlDataReaderPrimitiveKind.Double;
                GetValidatedOrdinalSlot(_currentDoubleValues, ordinal) = value;
                loaded = true;
                return true;
            }
        }

        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Parses a complete indexed numeric token through the common number
            /// parser. Formulas, missing values and XML-escaped text retain general
            /// value dispatch, which handles their additional semantics.
            /// </summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
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
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
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
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
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
