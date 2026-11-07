using System;
using System.Globalization;

namespace OfficeIMO.Spreadsheet {
    /// <summary>Width-independent numeric text used when spreadsheet formulas convert a number to text.</summary>
    internal static class SpreadsheetNumberText {
        internal static string FormatInvariant(double value) {
            if (value == 0d) return "0";
            string text = value.ToString("G15", CultureInfo.InvariantCulture);
            int separator = text.IndexOf('E');
            if (separator < 0) return text;

            string mantissa = text.Substring(0, separator);
            int exponent = int.Parse(text.Substring(separator + 1), CultureInfo.InvariantCulture);
            bool negative = mantissa[0] == '-';
            if (negative) mantissa = mantissa.Substring(1);
            int point = mantissa.IndexOf('.');
            int decimalPosition = (point < 0 ? mantissa.Length : point) + exponent;
            string digits = mantissa.Replace(".", string.Empty);
            int plainLength = decimalPosition <= 0 ? 2 - decimalPosition + digits.Length
                : decimalPosition >= digits.Length ? decimalPosition : digits.Length + 1;
            if (plainLength > 20) {
                return (negative ? "-" : string.Empty) + mantissa + "E"
                    + (exponent >= 0 ? "+" : "-") + Math.Abs(exponent).ToString(CultureInfo.InvariantCulture);
            }

            string plain = decimalPosition <= 0 ? "0." + new string('0', -decimalPosition) + digits
                : decimalPosition >= digits.Length ? digits + new string('0', decimalPosition - digits.Length)
                : digits.Insert(decimalPosition, ".");
            return negative ? "-" + plain : plain;
        }
    }
}
