using System.Globalization;

namespace OfficeIMO.Excel {
    internal static partial class ExcelNumberFormatDisplay {
        /// <summary>Formats the supported scientific mantissa while preserving optional precision and exponent presentation.</summary>
        private static bool TryFormatScientific(double value, string formatCode, int selectedSection, out string text) {
            text = string.Empty;
            int first = FindFirstNumericPlaceholder(formatCode);
            if (first < 0) return false;
            for (int index = first; index < formatCode.Length; index++) {
                char token = formatCode[index];
                if (token == LiteralPunctuationMarker && index + 1 < formatCode.Length) { index++; continue; }
                if (token is not 'E' and not 'e' || index + 2 >= formatCode.Length) continue;
                char sign = formatCode[index + 1];
                if (sign is not '+' and not '-') continue;
                int digits = 0;
                while (index + 2 + digits < formatCode.Length && formatCode[index + 2 + digits] == '0') digits++;
                if (digits == 0) continue;
                DecimalPlaceInfo precision = GetDecimalPlaceInfo(formatCode.Substring(first, index - first));
                string mantissa = "0" + (precision.Maximum == 0 ? string.Empty
                    : "." + new string('0', precision.Required) + new string('#', precision.Optional));
                string numericCode = mantissa + token + sign + new string('0', digits);
                double numericValue = selectedSection == 1 ? Math.Abs(value) : value;
                text = ApplyNumericAffixes(formatCode, numericValue.ToString(numericCode, CultureInfo.InvariantCulture));
                return true;
            }
            return false;
        }
    }
}
