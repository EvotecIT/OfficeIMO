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

            /// <summary>Retains IEEE-to-decimal semantics, the configured object type and lazy date values.</summary>
            private bool TryGetUnloadedDecimal(int ordinal, out decimal value) {
                EnsureOpenRow();
                if ((uint)ordinal >= (uint)_fieldCount) {
                    throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
                }
                value = default;
                if (_utf8Source == null || _currentRowIsBlank || _currentRow == null) return false;
                ref bool loaded = ref _currentValueLoaded[ordinal];
                if (loaded || _owner._opt.CellValueConverter != null) return false;
                bool numericAsDecimal = _owner._opt.NumericAsDecimal;
#if NET8_0_OR_GREATER
                bool tryLiteral = true;
#else
                bool tryLiteral = numericAsDecimal;
#endif
                if (tryLiteral && _utf8Source.TryGetShortDecimal(
                        ordinal + _utf8SourceOrdinalOffset, numericAsDecimal, out value, out double parsedNumber)) {
                    if (numericAsDecimal) {
                        GetValidatedOrdinalSlot(_currentPrimitiveKinds, ordinal) = XmlDataReaderPrimitiveKind.Decimal;
                        GetValidatedOrdinalSlot(_currentDecimalValues, ordinal) = value;
                    } else {
                        GetValidatedOrdinalSlot(_currentPrimitiveKinds, ordinal) = XmlDataReaderPrimitiveKind.Double;
                        GetValidatedOrdinalSlot(_currentDoubleValues, ordinal) = parsedNumber;
                    }
                    loaded = true;
                    return true;
                }
                if (!_utf8Source.TryGetNumber(ordinal + _utf8SourceOrdinalOffset, out double number, out bool dateStyle)
                    || !TryConvertExcelNumberToDecimal(number, out value)) return false;

                if (_owner._opt.NumericAsDecimal && !dateStyle) {
                    GetValidatedOrdinalSlot(_currentPrimitiveKinds, ordinal) = XmlDataReaderPrimitiveKind.Decimal;
                    GetValidatedOrdinalSlot(_currentDecimalValues, ordinal) = value;
                } else {
                    GetValidatedOrdinalSlot(_currentPrimitiveKinds, ordinal) = XmlDataReaderPrimitiveKind.Double;
                    GetValidatedOrdinalSlot(_currentDoubleValues, ordinal) = number;
                }
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
                if (loaded
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
            /// the configured numeric object type for subsequent scalar getters.
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
                if (loaded
                    || _owner._opt.CellValueConverter != null
                    || !_utf8Source.TryGetExactInt32(ordinal + _utf8SourceOrdinalOffset, out value)) {
                    return false;
                }

                if (_owner._opt.NumericAsDecimal) {
                    GetValidatedOrdinalSlot(_currentPrimitiveKinds, ordinal) = XmlDataReaderPrimitiveKind.Decimal;
                    GetValidatedOrdinalSlot(_currentDecimalValues, ordinal) = value;
                } else {
                    GetValidatedOrdinalSlot(_currentPrimitiveKinds, ordinal) = XmlDataReaderPrimitiveKind.Double;
                    GetValidatedOrdinalSlot(_currentDoubleValues, ordinal) = value;
                }
                loaded = true;
                return true;
            }
        }

        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Plain literals with at most fifteen digits fit the IEEE-to-decimal
            /// conversion's precision. Normalize their scale and zero just as that
            /// conversion does; other numeric spellings keep the ordinary parser.
            /// When rounding is disabled, also retain the ordinary double for the
            /// default object cache; a rounded decimal cannot reconstruct its bits.
            /// </summary>
            internal bool TryGetShortDecimal(int ordinal, bool allowRounded, out decimal value, out double number) {
                EnsureNotDisposed();
                value = default;
                number = default;
                int cellIndex = _currentRowOffset + ordinal;
                if (allowRounded && TryGetDecodedDecimal(cellIndex, out value)) return true;
                if (_cellKinds![cellIndex] != (byte)Utf8CellKind.Number
                    || _formulaLengths != null && _formulaLengths[cellIndex] >= 0
                    || _valueLengths![cellIndex] < 0) return false;

                ReadOnlySpan<byte> token = TrimAsciiWhitespace(
                    _buffer!.AsSpan(_valueStarts![cellIndex], _valueLengths[cellIndex]));
                if (token.IsEmpty) return false;
                bool negative = token[0] == (byte)'-';
                if (token.Length > (negative ? 17 : 16)) {
#if NET8_0_OR_GREATER
                    return allowRounded && ExcelUtf8NumberParser.TryParseRoundedFixedPointDecimal(token, out value);
#else
                    return false;
#endif
                }
                int index = negative ? 1 : 0;
                int digits = 0, scale = 0;
                bool point = false;
                ulong mantissa = 0;
                for (; index < token.Length; index++) {
                    uint digit = (uint)(token[index] - (byte)'0');
                    if (digit <= 9) {
                        if (++digits > 15) return false;
                        mantissa = mantissa * 10 + digit;
                        if (point) scale++;
                    } else if (token[index] == (byte)'.' && !point && digits > 0 && index < token.Length - 1) {
                        point = true;
                    } else return false;
                }
                if (digits == 0) return false;
#if NET8_0_OR_GREATER
                if (!allowRounded) {
                    // Use the same exact mantissa and division as the common double
                    // parser before normalizing the decimal coefficient and zero.
                    double parsed = mantissa;
                    if (scale != 0) parsed /= ExcelUtf8NumberParser.PowerOfTen(scale);
                    number = negative ? -parsed : parsed;
                }
#endif
                if (mantissa == 0) return true;
                while (scale > 0 && mantissa % 10 == 0) {
                    mantissa /= 10;
                    scale--;
                }
                value = new decimal(unchecked((int)mantissa), (int)(mantissa >> 32), 0, negative, (byte)scale);
                return true;
            }

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
                if (TryGetDecodedDouble(cellIndex, out value)) return true;
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
                if (TryGetDecodedDateTime(_currentRowOffset + ordinal, out value)) return true;
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
                if (TryGetDecodedInt32(cellIndex, out value)) return true;
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
