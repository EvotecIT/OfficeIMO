using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    internal static partial class ExcelNumberFormatDisplay {
        private static bool TryFormatFraction(double value, string formatCode, int selectedSection, out string text) {
            text = string.Empty;
            int slash = formatCode.IndexOf('/');
            if (slash < 0 || !HasFractionNumerator(formatCode, slash)) {
                return false;
            }

            bool fixedDenominator = TryGetFixedFractionDenominator(formatCode, slash, out int fixedValue);
            int maxDenominator = fixedDenominator
                ? fixedValue
                : GetFractionDenominatorLimit(formatCode, slash);
            if (maxDenominator <= 0) {
                return false;
            }

            double displayValue = selectedSection == 1 ? Math.Abs(value) : value;
            bool negative = displayValue < 0D;
            double absoluteValue = Math.Abs(displayValue);
            if (absoluteValue >= long.MaxValue) {
                // At this magnitude every finite double is an integer. Avoid
                // overflowing the mixed-fraction whole part or falling through
                // to decimal formatting that appends the denominator literally.
                text = displayValue.ToString("0", CultureInfo.InvariantCulture);
                return true;
            }

            long whole = (long)Math.Floor(absoluteValue);
            double fractional = absoluteValue - whole;
            long numerator = 0L;
            long denominator = 1L;

            if (fractional > 0.0000000001D) {
                if (fixedDenominator) {
                    denominator = fixedValue;
                    numerator = (int)Math.Round(fractional * denominator, MidpointRounding.AwayFromZero);
                } else {
                    double bestError = double.MaxValue;
                    for (int currentDenominator = 1; currentDenominator <= maxDenominator; currentDenominator++) {
                        long currentNumerator = (long)Math.Round(fractional * currentDenominator, MidpointRounding.AwayFromZero);
                        double candidate = currentNumerator / (double)currentDenominator;
                        double error = Math.Abs(candidate - fractional);
                        if (error + 0.0000000001D < bestError) {
                            bestError = error;
                            numerator = currentNumerator;
                            denominator = currentDenominator;
                        }
                    }
                }

                if (numerator >= denominator) {
                    if (whole > long.MaxValue - (numerator / denominator)) {
                        return false;
                    }

                    whole += numerator / denominator;
                    numerator %= denominator;
                }
            }

            bool mixedFraction = HasMixedFractionWholePart(formatCode, slash);
            if (!mixedFraction && numerator > 0) {
                if (whole > (long.MaxValue - numerator) / denominator) {
                    return false;
                }

                numerator += whole * denominator;
                whole = 0;
            }

            string numericText;
            if (numerator == 0) {
                numericText = whole.ToString(CultureInfo.InvariantCulture);
            } else if (whole > 0) {
                numericText = whole.ToString(CultureInfo.InvariantCulture) + " " +
                    numerator.ToString(CultureInfo.InvariantCulture) + "/" +
                    denominator.ToString(CultureInfo.InvariantCulture);
            } else {
                numericText = numerator.ToString(CultureInfo.InvariantCulture) + "/" +
                    denominator.ToString(CultureInfo.InvariantCulture);
            }

            text = negative && numericText != "0"
                ? "-" + numericText
                : numericText;
            return true;
        }

        private static bool HasFractionNumerator(string formatCode, int slash) {
            for (int i = slash - 1; i >= 0; i--) {
                char ch = formatCode[i];
                if (IsNumericPlaceholder(ch)) {
                    return true;
                }

                if (!char.IsWhiteSpace(ch)) {
                    continue;
                }
            }

            return false;
        }

        private static bool HasMixedFractionWholePart(string formatCode, int slash) {
            int lastSpace = formatCode.LastIndexOf(' ', Math.Max(0, slash - 1), slash);
            if (lastSpace <= 0) {
                return false;
            }

            for (int i = 0; i < lastSpace; i++) {
                if (IsNumericPlaceholder(formatCode[i])) {
                    return true;
                }
            }

            return false;
        }

        private static int GetFractionDenominatorLimit(string formatCode, int slash) {
            int places = 0;
            for (int i = slash + 1; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (IsNumericPlaceholder(ch)) {
                    places++;
                    continue;
                }

                if (places > 0) {
                    break;
                }
            }

            if (places <= 0) {
                return 0;
            }

            int limit = 1;
            for (int i = 0; i < Math.Min(places, 4); i++) {
                limit *= 10;
            }

            return limit - 1;
        }

        private static bool TryGetFixedFractionDenominator(string formatCode, int slash, out int denominator) {
            denominator = 0;
            var builder = new StringBuilder();
            for (int i = slash + 1; i < formatCode.Length; i++) {
                char ch = formatCode[i];
                if (char.IsDigit(ch)) {
                    builder.Append(ch);
                    continue;
                }

                if (builder.Length > 0) {
                    break;
                }

                if (!char.IsWhiteSpace(ch)) {
                    return false;
                }
            }

            return builder.Length > 0
                && int.TryParse(builder.ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out denominator)
                && denominator > 0;
        }

        private static string ApplyFractionAffixes(string formatCode, string numericText) {
            int first = FindFirstNumericPlaceholder(formatCode);
            int last = FindLastNumericPlaceholder(formatCode);
            int slash = formatCode.IndexOf('/');
            if (slash >= 0) {
                int denominatorEnd = slash;
                for (int i = slash + 1; i < formatCode.Length; i++) {
                    char ch = formatCode[i];
                    if (char.IsWhiteSpace(ch)) {
                        if (denominatorEnd > slash) {
                            break;
                        }

                        continue;
                    }

                    if (char.IsDigit(ch) || IsNumericPlaceholder(ch)) {
                        denominatorEnd = i;
                        continue;
                    }

                    break;
                }

                last = Math.Max(last, denominatorEnd);
            }

            if (first < 0 || last < first) {
                return CleanLiteralAffix(formatCode);
            }

            string prefix = CleanLiteralAffix(formatCode.Substring(0, first));
            string suffix = CleanLiteralAffix(formatCode.Substring(last + 1));
            return prefix + numericText + suffix;
        }

    }
}
