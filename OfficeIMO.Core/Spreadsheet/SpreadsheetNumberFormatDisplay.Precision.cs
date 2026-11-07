using System.Linq;
namespace OfficeIMO.Spreadsheet {
    internal static partial class SpreadsheetNumberFormatDisplay {
        private static DecimalPlaceInfo GetDecimalPlaceInfo(string formatCode) {
            int dot = -1;
            for (int index = 0; index < formatCode.Length; index++) {
                if (formatCode[index] == LiteralPunctuationMarker && index + 1 < formatCode.Length) { index++; continue; }
                if (formatCode[index] == '.') { dot = index; break; }
            }
            if (dot < 0) return new DecimalPlaceInfo(0, 0);
            int required = 0;
            int maximum = 0;
            for (int index = dot + 1; index < formatCode.Length; index++) {
                char token = formatCode[index];
                if (token == '0') { maximum++; required = maximum; continue; }
                if (token is '#' or '?') { maximum++; continue; }
                break;
            }
            return new DecimalPlaceInfo(required, maximum);
        }

        /// <summary>Counts the required integer digits ("0" placeholders before the decimal point).</summary>
        private static int GetRequiredIntegerDigits(string formatCode) {
            int required = 0;
            for (int index = 0; index < formatCode.Length; index++) {
                char token = formatCode[index];
                if (token == LiteralPunctuationMarker && index + 1 < formatCode.Length) { index++; continue; }
                if (token is '.' or 'e' or 'E' or '/') break;
                if (token == '0') required++;
            }
            return required;
        }

        /// <summary>Left-pads the integer digits so formats such as 000000 keep their leading zeros.</summary>
        private static string PadRequiredIntegerDigits(string text, int requiredDigits, bool thousands, System.Globalization.NumberFormatInfo numberFormat) {
            if (requiredDigits <= 1) return text;
            int start = text.StartsWith(numberFormat.NegativeSign, System.StringComparison.Ordinal) ? numberFormat.NegativeSign.Length : 0;
            int end = text.IndexOf(numberFormat.NumberDecimalSeparator, System.StringComparison.Ordinal);
            if (end < 0) end = text.Length;
            string integer = text.Substring(start, end - start);
            if (numberFormat.NumberGroupSeparator.Length > 0) integer = integer.Replace(numberFormat.NumberGroupSeparator, string.Empty);
            if (integer.Length >= requiredDigits || integer.Any(ch => ch < '0' || ch > '9')) return text;
            integer = integer.PadLeft(requiredDigits, '0');
            if (thousands) {
                var grouped = new System.Text.StringBuilder(integer.Length + integer.Length / 3);
                for (int index = 0; index < integer.Length; index++) {
                    if (index > 0 && (integer.Length - index) % 3 == 0) grouped.Append(numberFormat.NumberGroupSeparator);
                    grouped.Append(integer[index]);
                }
                integer = grouped.ToString();
            }
            return text.Substring(0, start) + integer + text.Substring(end);
        }

        private static string TrimOptionalDecimalPlaces(string text, int requiredDecimalPlaces, System.Globalization.NumberFormatInfo numberFormat) {
            int dot = text.IndexOf(numberFormat.NumberDecimalSeparator, System.StringComparison.Ordinal);
            if (dot < 0) return text;
            int fractionStart = dot + numberFormat.NumberDecimalSeparator.Length;
            int end = text.Length - 1;
            while (end >= fractionStart + requiredDecimalPlaces && text[end] == '0') end--;
            if (end < fractionStart) return text.Substring(0, dot);
            return end == text.Length - 1 ? text : text.Substring(0, end + 1);
        }

        private readonly struct DecimalPlaceInfo {
            internal DecimalPlaceInfo(int required, int maximum) {
                Required = required;
                Maximum = maximum;
            }
            internal int Required { get; }
            internal int Maximum { get; }
            internal int Optional => Maximum - Required;
        }
    }
}
