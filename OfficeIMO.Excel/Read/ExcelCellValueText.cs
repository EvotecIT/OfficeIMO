using System.Globalization;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel {
    /// <summary>
    /// The text a data reader returns for a typed cell value. Numbers use the shortest text that parses back to the
    /// same value on every target; .NET Framework's default number text keeps only 15 digits.
    /// </summary>
    internal static class ExcelCellValueText {
        internal static string Format(object? value, CultureInfo culture) => value switch {
            double number => OfficeRoundTripText.Format(number, culture),
            float number => OfficeRoundTripText.Format(number, culture),
            _ => Convert.ToString(value, culture) ?? string.Empty
        };
    }
}
