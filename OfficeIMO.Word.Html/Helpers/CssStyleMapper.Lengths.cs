using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal static partial class CssStyleMapper {
        private static bool TryParseLength(
            string value,
            double rootFontSizePixels,
            double elementFontSizePixels,
            out int twips) {
            twips = 0;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            string normalized = value.Trim().ToLowerInvariant();
            if (TryParseLengthUnit(normalized, "rem", rootFontSizePixels * 15d, out twips) ||
                TryParseLengthUnit(normalized, "em", elementFontSizePixels * 15d, out twips) ||
                TryParseLengthUnit(normalized, "px", 15d, out twips) ||
                TryParseLengthUnit(normalized, "pt", 20d, out twips) ||
                TryParseLengthUnit(normalized, "pc", 240d, out twips) ||
                TryParseLengthUnit(normalized, "in", 1440d, out twips) ||
                TryParseLengthUnit(normalized, "cm", 1440d / 2.54d, out twips) ||
                TryParseLengthUnit(normalized, "mm", 1440d / 25.4d, out twips) ||
                TryParseLengthUnit(normalized, "q", 1440d / 101.6d, out twips)) {
                return true;
            }

            if (double.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out double number)) {
                twips = (int)Math.Round(number * 15d);
                return true;
            }

            return false;
        }

        private static bool TryParseLengthUnit(
            string value,
            string unit,
            double twipsPerUnit,
            out int twips) {
            twips = 0;
            if (!value.EndsWith(unit, StringComparison.Ordinal) ||
                !double.TryParse(
                    value.Substring(0, value.Length - unit.Length),
                    NumberStyles.Number,
                    CultureInfo.InvariantCulture,
                    out double number)) {
                return false;
            }

            twips = (int)Math.Round(number * twipsPerUnit);
            return true;
        }
    }
}
