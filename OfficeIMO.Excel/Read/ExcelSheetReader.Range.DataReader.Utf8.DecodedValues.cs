#nullable enable

using System.Buffers;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private const byte DecodedDecimal = 1;
            private const byte DecodedDouble = 2;
            private const byte DecodedDate = 3;
            private const byte DecodedKindMask = 0x0F;
            private const byte DecodedExactInt32 = 0x10;
            private const int DecodedMantissaBits = 50;
            private const ulong DecodedMantissaMask = (1UL << DecodedMantissaBits) - 1;

            // Parallel indexing helpers decode numeric cells ahead of delivery. Each slot holds
            // the exact result of the ordinary getter helper for that cell; a zero tag means the
            // getter parses the indexed token itself.
            private long[]? _decodedValues;
            private byte[]? _decodedTags;
            private int _decodedThroughCell;

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private bool TryGetDecodedDecimal(int cellIndex, out decimal value) {
                byte[]? tags = _decodedTags;
                if (tags == null || (tags[cellIndex] & DecodedKindMask) != DecodedDecimal) {
                    value = default;
                    return false;
                }
                value = UnpackDecimal(_decodedValues![cellIndex]);
                return true;
            }

            /// <summary>Reads a decoded current-row decimal without entering the literal parser.</summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            internal bool TryGetDecodedDecimalAt(int ordinal, out decimal value) =>
                TryGetDecodedDecimal(_currentRowOffset + ordinal, out value);

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private bool TryGetDecodedDouble(int cellIndex, out double value) {
                byte[]? tags = _decodedTags;
                if (tags == null || (tags[cellIndex] & DecodedKindMask) != DecodedDouble) {
                    value = default;
                    return false;
                }
                value = BitConverter.Int64BitsToDouble(_decodedValues![cellIndex]);
                return true;
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private bool TryGetDecodedDateTime(int cellIndex, out DateTime value) {
                byte[]? tags = _decodedTags;
                if (tags == null || (tags[cellIndex] & DecodedKindMask) != DecodedDate) {
                    value = default;
                    return false;
                }
                value = new DateTime(_decodedValues![cellIndex], DateTimeKind.Unspecified);
                return true;
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private bool TryGetDecodedInt32(int cellIndex, out int value) {
                byte[]? tags = _decodedTags;
                byte tag = tags == null ? (byte)0 : tags[cellIndex];
                if ((tag & DecodedExactInt32) == 0) {
                    value = default;
                    return false;
                }
                long slot = _decodedValues![cellIndex];
                value = (tag & DecodedKindMask) == DecodedDecimal
                    ? UnpackDecimalInteger(slot)
                    : (int)BitConverter.Int64BitsToDouble(slot);
                return true;
            }

            /// <summary>Decodes one indexed row through the same helpers the getters use.</summary>
            private void DecodeRowValues(int rowOffset, long[] values, byte[] tags) {
                int savedOffset = _currentRowOffset;
                _currentRowOffset = rowOffset;
                bool asDecimal = _options.NumericAsDecimal;
                try {
                    for (int ordinal = 0; ordinal < _fieldCount; ordinal++) {
                        int cell = rowOffset + ordinal;
                        byte tag = 0;
                        long slot = 0;
                        byte kind = _cellKinds![cell];
                        if ((kind & CellKindMask) == (byte)Utf8CellKind.Number && _valueLengths![cell] >= 0) {
                            if ((kind & DateStyleCellKindFlag) != 0) {
                                if (TryDecodeDateTime(ordinal, out DateTime date)) {
                                    tag = DecodedDate;
                                    slot = date.Ticks;
                                }
                            } else if (asDecimal) {
                                if (TryGetShortDecimal(ordinal, allowRounded: true, out decimal value, out _)
                                    && TryPackDecimal(value, out slot)) {
                                    tag = DecodedDecimal;
                                    if (TryGetExactInt32(ordinal, out int integer) && IsPackedInteger(slot, integer)) {
                                        tag |= DecodedExactInt32;
                                    }
                                }
                            } else if (TryGetNumber(ordinal, out double number, out _)) {
                                tag = DecodedDouble;
                                slot = BitConverter.DoubleToInt64Bits(number);
                                if (TryGetExactInt32(ordinal, out int integer) && integer == number) {
                                    tag |= DecodedExactInt32;
                                }
                            }
                        }
                        tags[cell] = tag;
                        values[cell] = slot;
                    }
                } finally {
                    _currentRowOffset = savedOffset;
                }
            }

            // Out-of-range serials stay undecoded so the getter raises its ordinary conversion error.
            private bool TryDecodeDateTime(int ordinal, out DateTime date) {
                try {
                    return TryGetDateTime(ordinal, out date) && date.Kind == DateTimeKind.Unspecified;
                } catch (ArgumentException) {
                    date = default;
                    return false;
                }
            }

            private static bool TryPackDecimal(decimal value, out long packed) {
                packed = 0;
#if NET8_0_OR_GREATER
                Span<int> bits = stackalloc int[4];
                decimal.GetBits(value, bits);
#else
                int[] bits = decimal.GetBits(value);
#endif
                ulong mantissa = (uint)bits[0] | ((ulong)(uint)bits[1] << 32);
                int scale = (bits[3] >> 16) & 0xFF;
                if (bits[2] != 0 || mantissa > DecodedMantissaMask || scale > 31) return false;
                packed = (long)(mantissa | ((ulong)scale << DecodedMantissaBits) | (bits[3] < 0 ? 1UL << 55 : 0));
                return true;
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static decimal UnpackDecimal(long packed) {
                ulong bits = (ulong)packed;
                ulong mantissa = bits & DecodedMantissaMask;
                return new decimal(unchecked((int)mantissa), (int)(mantissa >> 32), 0,
                    (bits & (1UL << 55)) != 0, (byte)((bits >> DecodedMantissaBits) & 31));
            }

            private static bool IsPackedInteger(long packed, int integer) {
                ulong bits = (ulong)packed;
                if (((bits >> DecodedMantissaBits) & 31) != 0) return false;
                long mantissa = (long)(bits & DecodedMantissaMask);
                long signed = (bits & (1UL << 55)) != 0 ? -mantissa : mantissa;
                return signed == integer;
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static int UnpackDecimalInteger(long packed) {
                ulong bits = (ulong)packed;
                long mantissa = (long)(bits & DecodedMantissaMask);
                return (int)((bits & (1UL << 55)) != 0 ? -mantissa : mantissa);
            }

            /// <summary>Allocates the decoded lane when the first helper chunk is adopted.</summary>
            private void EnsureDecodedValues() {
                if (_decodedTags != null) return;
                int capacity = _valueStarts!.Length;
                _decodedValues = ExcelArrayPool.RentLongs(capacity);
                _decodedTags = ExcelArrayPool.RentBytes(capacity);
                _decodedThroughCell = 0;
            }

            /// <summary>Marks cells the sequential indexer handled itself as undecoded.</summary>
            private void ClearDecodedValuesThrough(int cellEnd) {
                if (_decodedTags == null || cellEnd <= _decodedThroughCell) return;
                Array.Clear(_decodedTags, _decodedThroughCell, cellEnd - _decodedThroughCell);
                _decodedThroughCell = cellEnd;
            }

            private void GrowDecodedValues(int capacity, int count) {
                if (_decodedTags == null) return;
                long[] values = ExcelArrayPool.RentLongs(capacity);
                byte[] tags = ExcelArrayPool.RentBytes(capacity);
                Array.Copy(_decodedValues!, values, count);
                Array.Copy(_decodedTags, tags, count);
                ExcelArrayPool.Return(_decodedValues!);
                ExcelArrayPool.Return(_decodedTags);
                _decodedValues = values;
                _decodedTags = tags;
            }

            private void ReturnDecodedValues() {
                if (_decodedValues != null) ExcelArrayPool.Return(_decodedValues);
                if (_decodedTags != null) ExcelArrayPool.Return(_decodedTags);
                _decodedValues = null;
                _decodedTags = null;
                _decodedThroughCell = 0;
            }
        }
    }
}
