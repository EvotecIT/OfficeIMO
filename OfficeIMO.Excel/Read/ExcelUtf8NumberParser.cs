#nullable enable

using System.Buffers.Text;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Parses invariant UTF-8 numbers emitted by spreadsheet formats while preserving the
    /// general-purpose parser as a fallback for uncommon or high-precision representations.
    /// </summary>
    internal static class ExcelUtf8NumberParser {
        /// <summary>
        /// Parses a complete invariant UTF-8 double value.
        /// </summary>
        internal static bool TryParseDouble(ReadOnlySpan<byte> value, out double result) {
#if NET8_0_OR_GREATER
            return TryParseCommonFixedPoint(value, out result)
                || (Utf8Parser.TryParse(value, out result, out int consumed) && consumed == value.Length);
#else
            return Utf8Parser.TryParse(value, out result, out int consumed) && consumed == value.Length;
#endif
        }

#if NET8_0_OR_GREATER
        /// <summary>
        /// Parses the common fixed-point form without entering the general-purpose floating-point
        /// parser. The conservative bounds keep the integer mantissa exactly representable.
        /// </summary>
        internal static bool TryParseCommonFixedPoint(ReadOnlySpan<byte> value, out double result) {
            result = 0;
            if (value.IsEmpty) {
                return false;
            }

            int index = 0;
            bool negative = value[0] == (byte)'-';
            if (negative || value[0] == (byte)'+') {
                index++;
                if (index == value.Length) {
                    return false;
                }
            }

            ulong mantissa = 0;
            int significantDigits = 0;
            int fractionalDigits = 0;
            bool hasDigit = false;
            bool hasDecimalPoint = false;
            for (; index < value.Length; index++) {
                byte current = value[index];
                if (current == (byte)'.') {
                    if (hasDecimalPoint) {
                        return false;
                    }

                    hasDecimalPoint = true;
                    continue;
                }

                uint digit = (uint)(current - (byte)'0');
                if (digit > 9) {
                    return false;
                }

                hasDigit = true;
                if (mantissa != 0 || digit != 0) {
                    significantDigits++;
                    if (significantDigits > 15) {
                        return false;
                    }
                }

                mantissa = (mantissa * 10) + digit;
                if (hasDecimalPoint && ++fractionalDigits > 22) {
                    return false;
                }
            }

            if (!hasDigit) {
                return false;
            }

            double parsed = mantissa;
            if (fractionalDigits != 0) {
                parsed /= PowerOfTen(fractionalDigits);
            }

            result = negative ? -parsed : parsed;
            return true;
        }

        private static double PowerOfTen(int exponent) => exponent switch {
            1 => 1e1,
            2 => 1e2,
            3 => 1e3,
            4 => 1e4,
            5 => 1e5,
            6 => 1e6,
            7 => 1e7,
            8 => 1e8,
            9 => 1e9,
            10 => 1e10,
            11 => 1e11,
            12 => 1e12,
            13 => 1e13,
            14 => 1e14,
            15 => 1e15,
            16 => 1e16,
            17 => 1e17,
            18 => 1e18,
            19 => 1e19,
            20 => 1e20,
            21 => 1e21,
            _ => 1e22
        };
#endif
    }
}
