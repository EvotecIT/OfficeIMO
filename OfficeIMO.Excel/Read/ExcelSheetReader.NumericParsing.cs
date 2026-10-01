using System.Globalization;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private static bool TryParseInvariantDouble(string? rawText, out double value) {
            value = 0;
            if (string.IsNullOrEmpty(rawText)) {
                return false;
            }

            return TryParseInvariantDouble(rawText.AsSpan(), out value);
        }

        // Use the runtime parser's correctly rounded conversion. Accumulating
        // fractional digits with repeated double additions can cross a format
        // midpoint or change a cached value by one or more ULPs.
        private static bool TryParseInvariantDouble(ReadOnlySpan<char> text, out double value) {
#if NETSTANDARD2_0
            return double.TryParse(text.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
#else
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
#endif
        }

        private bool TryParseRawDouble(string rawText, out double value) {
            if (_opt.Culture != CultureInfo.InvariantCulture
                && double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out value)) {
                return true;
            }

            return TryParseInvariantDouble(rawText, out value)
                || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);
        }

        private bool TryParseRawInt32(string rawText, out int value) {
            if (_opt.Culture == CultureInfo.InvariantCulture && TryParseInvariantInt32Fast(rawText, out value)) {
                return true;
            }

            return int.TryParse(rawText, NumberStyles.Integer, _opt.Culture, out value);
        }

        private static bool TryParseInvariantInt32Fast(string rawText, out int value) {
            value = 0;
            if (string.IsNullOrEmpty(rawText)) {
                return false;
            }

            int index = 0;
            bool negative = false;
            if (rawText[0] == '-') {
                negative = true;
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            } else if (rawText[0] == '+') {
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            }

            uint limit = negative ? 2147483648U : int.MaxValue;
            uint parsed = 0U;
            for (; index < rawText.Length; index++) {
                int digit = rawText[index] - '0';
                if ((uint)digit > 9U) {
                    return false;
                }

                if (parsed > (limit - (uint)digit) / 10U) {
                    return false;
                }

                parsed = (parsed * 10U) + (uint)digit;
            }

            if (negative) {
                value = parsed == 2147483648U ? int.MinValue : -(int)parsed;
            } else {
                value = (int)parsed;
            }

            return true;
        }

        private bool TryParseRawInt64(string rawText, out long value) {
            if (_opt.Culture == CultureInfo.InvariantCulture && TryParseInvariantInt64Fast(rawText, out value)) {
                return true;
            }

            return long.TryParse(rawText, NumberStyles.Integer, _opt.Culture, out value);
        }

        private static bool TryParseInvariantInt64Fast(string rawText, out long value) {
            value = 0;
            if (string.IsNullOrEmpty(rawText)) {
                return false;
            }

            int index = 0;
            bool negative = false;
            if (rawText[0] == '-') {
                negative = true;
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            } else if (rawText[0] == '+') {
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            }

            ulong limit = negative ? 9223372036854775808UL : long.MaxValue;
            ulong parsed = 0UL;
            for (; index < rawText.Length; index++) {
                int digit = rawText[index] - '0';
                if ((uint)digit > 9U) {
                    return false;
                }

                if (parsed > (limit - (uint)digit) / 10UL) {
                    return false;
                }

                parsed = (parsed * 10UL) + (uint)digit;
            }

            if (negative) {
                value = parsed == 9223372036854775808UL ? long.MinValue : -(long)parsed;
            } else {
                value = (long)parsed;
            }

            return true;
        }

        private bool TryParseRawDecimal(string rawText, out decimal value) {
            return TryParseRawDecimal(rawText, _opt.Culture, out value);
        }

        private static bool TryParseRawDecimal(string rawText, CultureInfo culture, out decimal value) {
            if (culture == CultureInfo.InvariantCulture) {
                return TryParseInvariantDecimalFast(rawText, out value)
                    || decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);
            }

            if (decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, culture, out value)) {
                return true;
            }

            return TryParseInvariantDecimalFast(rawText, out value)
                || decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);
        }

        private static bool TryParseExcelNumberAsDecimal(string rawText, CultureInfo culture, out decimal value) {
            value = 0m;
            bool parsed = TryParseInvariantDouble(rawText, out double number)
                || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number)
                || (culture != CultureInfo.InvariantCulture
                    && double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, culture, out number));
            return parsed && TryConvertExcelNumberToDecimal(number, out value);
        }

        private static bool TryConvertExcelNumberToDecimal(double number, out decimal value) {
            value = 0m;
            if (double.IsNaN(number) || double.IsInfinity(number)) {
                return false;
            }

            try {
                value = (decimal)number;
                return true;
            } catch (OverflowException) {
                return false;
            }
        }

        private static bool TryParseInvariantDecimalFast(string rawText, out decimal value) {
            value = 0m;
            if (string.IsNullOrEmpty(rawText)) {
                return false;
            }

            int index = 0;
            bool negative = false;
            if (rawText[0] == '-') {
                negative = true;
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            } else if (rawText[0] == '+') {
                index = 1;
                if (index == rawText.Length) {
                    return false;
                }
            }

            ulong parsed = 0UL;
            int scale = 0;
            bool hasDigit = false;
            bool hasDecimalPoint = false;
            for (; index < rawText.Length; index++) {
                char ch = rawText[index];
                int digit = ch - '0';
                if ((uint)digit <= 9U) {
                    if (parsed > (ulong.MaxValue - (uint)digit) / 10UL) {
                        return false;
                    }

                    parsed = (parsed * 10UL) + (uint)digit;
                    hasDigit = true;
                    if (hasDecimalPoint && ++scale > 28) {
                        return false;
                    }

                    continue;
                }

                if (ch == '.' && !hasDecimalPoint) {
                    hasDecimalPoint = true;
                    continue;
                }

                return false;
            }

            if (!hasDigit) {
                return false;
            }

            value = new decimal(
                (int)(parsed & 0xFFFFFFFF),
                (int)((parsed >> 32) & 0xFFFFFFFF),
                0,
                negative,
                (byte)scale);
            return true;
        }

    }
}
