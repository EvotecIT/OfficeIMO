using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Spreadsheet;

internal static partial class SpreadsheetNumberFormatDisplay {
    /// <summary>Localizes canonical width-independent number text before any format literals are composed.</summary>
    internal static string FormatGeneralValue(double value, IFormatProvider provider) {
        string text = SpreadsheetNumberText.FormatInvariant(value);
        var numberFormat = NumberFormatInfo.GetInstance(provider);
        if (numberFormat.NumberDecimalSeparator == "." && numberFormat.NegativeSign == "-" && numberFormat.PositiveSign == "+") return text;
        var builder = new StringBuilder(text.Length);
        foreach (char ch in text) {
            if (ch == '.') builder.Append(numberFormat.NumberDecimalSeparator);
            else if (ch == '-') builder.Append(numberFormat.NegativeSign);
            else if (ch == '+') builder.Append(numberFormat.PositiveSign);
            else builder.Append(ch);
        }
        return builder.ToString();
    }
}
