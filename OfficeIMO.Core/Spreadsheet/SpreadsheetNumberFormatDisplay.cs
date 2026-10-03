using System;
using System.Globalization;

namespace OfficeIMO.Spreadsheet {
    internal static partial class SpreadsheetNumberFormatDisplay {
        internal const char LiteralPunctuationMarker = '\u0001';

        /// <summary>Formats invariant numeric cell text; null preserves the caller's raw fallback when the value or scaled result cannot be represented.</summary>
        internal static string? FormatNumericValue(double value, string formatCode) {
            if (double.IsNaN(value) || double.IsInfinity(value)) return null;
            int preferredSection = value < 0 ? 1 : value == 0 ? 2 : 0;
            string section = SelectNumberFormatSection(formatCode, preferredSection, value, out int selectedSection);
            if (section.Length == 0) {
                return string.Empty;
            }

            string normalized = StripNumberFormatDecorations(section);
            string lower = normalized.ToLowerInvariant();

            if (lower.Contains("@")) {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            if (!ContainsNumericPlaceholder(normalized)) {
                string literal = CleanLiteralAffix(normalized);
                return string.IsNullOrEmpty(literal) ? null : literal;
            }

            if (IsZeroValue(value) && HasOnlyOptionalDigitPlaceholders(normalized)) {
                return string.Empty;
            }

            if (TryFormatFraction(value, normalized, selectedSection, out string fractionText)) {
                return ApplyFractionAffixes(normalized, fractionText);
            }

            if (TryFormatScientific(value, normalized, selectedSection, out string scientificText)) return scientificText;

            int percentPlaceholders = CountPercentPlaceholders(section);
            bool thousands = lower.Contains("#,##") || lower.Contains(",##");
            DecimalPlaceInfo decimalPlaces = GetDecimalPlaceInfo(lower);
            double displayValue = value;
            for (int i = 0; i < percentPlaceholders; i++) {
                displayValue *= 100D;
            }

            int scalingCommas = CountScalingCommas(normalized);
            for (int i = 0; i < scalingCommas; i++) {
                displayValue /= 1000D;
            }

            if (double.IsNaN(displayValue) || double.IsInfinity(displayValue)) return null;

            if (selectedSection == 1) {
                displayValue = Math.Abs(displayValue);
            }

            string numericFormat = thousands
                ? "N" + decimalPlaces.Maximum.ToString(CultureInfo.InvariantCulture)
                : "F" + decimalPlaces.Maximum.ToString(CultureInfo.InvariantCulture);
            double roundedValue = decimalPlaces.Maximum <= 15
                && !double.IsNaN(displayValue) && !double.IsInfinity(displayValue)
                ? Math.Round(displayValue, decimalPlaces.Maximum, MidpointRounding.AwayFromZero)
                : displayValue;
            string text = roundedValue.ToString(numericFormat, CultureInfo.InvariantCulture);
            if (decimalPlaces.Optional > 0) {
                text = TrimOptionalDecimalPlaces(text, decimalPlaces.Required);
            }

            return ApplyNumericAffixes(normalized, text);
        }

    }
}
